"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { supabase } from "../lib/supabase";

function getElaboratoKey(r: any) {
  return (
    r.elaborato ||
    r.codiceElaborato ||
    r.codice_elaborato ||
    r.codice ||
    r.titolo ||
    r.id ||
    ""
  );
}

// Chiave tecnica per conteggiare gli elaborati una sola volta.
// Esempio: "T1266 0000 PE OS GTA..." e "T1266-0000-PE-OS-GTA..." diventano uguali.
function normalizeElaboratoCode(v = "") {
  return String(v || "")
    .toUpperCase()
    .replace(/\.PDF$/i, "")
    .replace(/[^A-Z0-9]/g, "");
}

function getElaboratoNormalizedKey(r: any) {
  const raw = getElaboratoKey(r);
  const normalized = normalizeElaboratoCode(raw);
  return normalized || cleanPdfText(raw);
}

function getElaboratoDisplay(r: any) {
  return cleanPdfText(getElaboratoKey(r)) || "Elaborato non identificato";
}

function translateStatus(status = "") {
  const value = String(status || "").trim();

  if (value === "Closed") return "Chiusa";
  if (value === "New") return "Aperta";
  if (value === "Waiting") return "In attesa";
  if (value === "Unknown") return "Non definito";

  return value || "Non definito";
}

function commentsToText(comments: any[]) {
  if (!Array.isArray(comments) || comments.length === 0) return "";

  return comments
    .map((c: any) => {
      const role = c.role ? `[${c.role}] ` : "";
      const author = c.author || "Autore non indicato";
      const date = c.date ? ` - ${c.date}` : "";
      const comment = c.comment || "";
      return `${role}${author}${date}\n${comment}`;
    })
    .join("\n\n");
}

function exportExcel(nomeFile: string, dati: any[]) {
  if (!dati || dati.length === 0) return;

  const ws = XLSX.utils.json_to_sheet(dati);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "Export");
  XLSX.writeFile(wb, `${nomeFile}.xlsx`);
}

function cleanPdfText(value: any) {
  return String(value ?? "")
    .replace(/[\u0000-\u001f\u007f]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}



function splitElaboratiValue(value: any) {
  const raw = cleanPdfText(value);
  if (!raw) return [];

  const normalized = raw
    .replace(/\s*\|\s*/g, "\n")
    .replace(/\s*;\s*/g, "\n")
    .replace(/\s+\/\s+/g, "\n");

  return Array.from(
    new Set(
      normalized
        .split(/\n+/)
        .map((item) => cleanPdfText(item))
        .filter(Boolean)
    )
  );
}

function getElaboratiForExportRow(row: any) {
  const raw = getElaboratoKey(row);
  const split = splitElaboratiValue(raw);
  return split.length > 0 ? split : [cleanPdfText(raw)].filter(Boolean);
}

function isGeneralSummaryRow(row: any) {
  const elaborato = cleanPdfText(getElaboratoKey(row));
  const normalized = normalizeElaboratoCode(elaborato);
  const descrizione = cleanPdfText(row?.descrizione || row?.title || row?.titolo || "").toLowerCase();

  if (!elaborato || elaborato === "Elaborato non identificato") return true;
  if (normalized.includes("GENERAL") || normalized.includes("GENERALE") || normalized.includes("GENERALI")) return true;
  if (descrizione.includes("rilievi generali") || descrizione.includes("osservazioni generali")) return true;

  return false;
}

function prepareRowsForPdfExport(rows: any[]) {
  const result: any[] = [];

  (rows || []).forEach((row: any) => {
    const elaborati = getElaboratiForExportRow(row);

    if (elaborati.length <= 1) {
      result.push(row);
      return;
    }

    elaborati.forEach((elaborato) => {
      result.push({
        ...row,
        elaborato,
        elaboratoDisplay: elaborato,
        elaboratoKey: normalizeElaboratoCode(elaborato),
      });
    });
  });

  return result;
}

function buildElaboratiDisciplinaSummary(rows: any[]) {
  const grouped: Record<string, { elaborato: string; nc: number; oss: number; nessun: number }> = {};

  rows.forEach((r: any) => {
    if (isGeneralSummaryRow(r)) return;

    const tipo = cleanPdfText(r.tipo || r["Tipologia rilievo"]).toUpperCase();
    const elaborati = getElaboratiForExportRow(r);

    elaborati.forEach((elaboratoRaw) => {
      const elaborato = cleanPdfText(elaboratoRaw) || "Elaborato non identificato";
      const key = normalizeElaboratoCode(elaborato) || elaborato;

      if (!grouped[key]) {
        grouped[key] = { elaborato, nc: 0, oss: 0, nessun: 0 };
      }

      if (tipo === "NC") {
        grouped[key].nc += 1;
      }

      // Le righe "Da NC a OSS" vanno conteggiate come OSS.
      if (tipo === "OSS" || tipo === "DA NC A OSS") {
        grouped[key].oss += 1;
      }

      if (tipo === "NESSUN RILIEVO") {
        grouped[key].nessun += 1;
      }
    });
  });

  return Object.values(grouped)
    .sort((a, b) => a.elaborato.localeCompare(b.elaborato, "it"))
    .map((item, index) => {
      const esito = item.nc === 0 && item.oss === 0
        ? "Nessuna NC/OSS"
        : `NC: ${item.nc} - OSS: ${item.oss}`;

      return [
        String(index + 1),
        item.elaborato,
        String(item.nc),
        String(item.oss),
        esito,
      ];
    });
}


type PdfHeaderData = {
  committente?: string;
  oggetto?: string;
  responsabileTecnico?: string;
  ispettore?: string;
  firma?: string;
  firmaImage?: string;
  firmaResponsabileImage?: string;
  firmaIspettoreImages?: string[];
  ispettoreKeys?: string[];
  responsabileTecnicoKey?: string;
  codiceScheda?: string;
  notaRicezione?: string;
  dataRicezione?: string;
  dataEmissione?: string;
};

const ITS_BLUE = [30, 117, 192] as [number, number, number];
const ITS_DARK_BLUE = [26, 92, 150] as [number, number, number];
const ITS_LIGHT_BLUE = [217, 232, 245] as [number, number, number];
const ITS_TABLE_LIGHT = [234, 243, 251] as [number, number, number];
const ITS_GRAY = [245, 245, 245] as [number, number, number];
const ITS_LOGO_BASE64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAOYAAABcCAIAAAD1WV7/AAAKMWlDQ1BJQ0MgUHJvZmlsZQAAeJydlndUU9kWh8+9N71QkhCKlNBraFICSA29SJEuKjEJEErAkAAiNkRUcERRkaYIMijggKNDkbEiioUBUbHrBBlE1HFwFBuWSWStGd+8ee/Nm98f935rn73P3Wfvfda6AJD8gwXCTFgJgAyhWBTh58WIjYtnYAcBDPAAA2wA4HCzs0IW+EYCmQJ82IxsmRP4F726DiD5+yrTP4zBAP+flLlZIjEAUJiM5/L42VwZF8k4PVecJbdPyZi2NE3OMErOIlmCMlaTc/IsW3z2mWUPOfMyhDwZy3PO4mXw5Nwn4405Er6MkWAZF+cI+LkyviZjg3RJhkDGb+SxGXxONgAoktwu5nNTZGwtY5IoMoIt43kA4EjJX/DSL1jMzxPLD8XOzFouEiSniBkmXFOGjZMTi+HPz03ni8XMMA43jSPiMdiZGVkc4XIAZs/8WRR5bRmyIjvYODk4MG0tbb4o1H9d/JuS93aWXoR/7hlEH/jD9ld+mQ0AsKZltdn6h21pFQBd6wFQu/2HzWAvAIqyvnUOfXEeunxeUsTiLGcrq9zcXEsBn2spL+jv+p8Of0NffM9Svt3v5WF485M4knQxQ143bmZ6pkTEyM7icPkM5p+H+B8H/nUeFhH8JL6IL5RFRMumTCBMlrVbyBOIBZlChkD4n5r4D8P+pNm5lona+BHQllgCpSEaQH4eACgqESAJe2Qr0O99C8ZHA/nNi9GZmJ37z4L+fVe4TP7IFiR/jmNHRDK4ElHO7Jr8WgI0IABFQAPqQBvoAxPABLbAEbgAD+ADAkEoiARxYDHgghSQAUQgFxSAtaAYlIKtYCeoBnWgETSDNnAYdIFj4DQ4By6By2AE3AFSMA6egCnwCsxAEISFyBAVUod0IEPIHLKFWJAb5AMFQxFQHJQIJUNCSAIVQOugUqgcqobqoWboW+godBq6AA1Dt6BRaBL6FXoHIzAJpsFasBFsBbNgTzgIjoQXwcnwMjgfLoK3wJVwA3wQ7oRPw5fgEVgKP4GnEYAQETqiizARFsJGQpF4JAkRIauQEqQCaUDakB6kH7mKSJGnyFsUBkVFMVBMlAvKHxWF4qKWoVahNqOqUQdQnag+1FXUKGoK9RFNRmuizdHO6AB0LDoZnYsuRlegm9Ad6LPoEfQ4+hUGg6FjjDGOGH9MHCYVswKzGbMb0445hRnGjGGmsVisOtYc64oNxXKwYmwxtgp7EHsSewU7jn2DI+J0cLY4X1w8TogrxFXgWnAncFdwE7gZvBLeEO+MD8Xz8MvxZfhGfA9+CD+OnyEoE4wJroRIQiphLaGS0EY4S7hLeEEkEvWITsRwooC4hlhJPEQ8TxwlviVRSGYkNimBJCFtIe0nnSLdIr0gk8lGZA9yPFlM3kJuJp8h3ye/UaAqWCoEKPAUVivUKHQqXFF4pohXNFT0VFysmK9YoXhEcUjxqRJeyUiJrcRRWqVUo3RU6YbStDJV2UY5VDlDebNyi/IF5UcULMWI4kPhUYoo+yhnKGNUhKpPZVO51HXURupZ6jgNQzOmBdBSaaW0b2iDtCkVioqdSrRKnkqNynEVKR2hG9ED6On0Mvph+nX6O1UtVU9Vvuom1TbVK6qv1eaoeajx1UrU2tVG1N6pM9R91NPUt6l3qd/TQGmYaYRr5Grs0Tir8XQObY7LHO6ckjmH59zWhDXNNCM0V2ju0xzQnNbS1vLTytKq0jqj9VSbru2hnaq9Q/uE9qQOVcdNR6CzQ+ekzmOGCsOTkc6oZPQxpnQ1df11Jbr1uoO6M3rGelF6hXrtevf0Cfos/ST9Hfq9+lMGOgYhBgUGrQa3DfGGLMMUw12G/YavjYyNYow2GHUZPTJWMw4wzjduNb5rQjZxN1lm0mByzRRjyjJNM91tetkMNrM3SzGrMRsyh80dzAXmu82HLdAWThZCiwaLG0wS05OZw2xljlrSLYMtCy27LJ9ZGVjFW22z6rf6aG1vnW7daH3HhmITaFNo02Pzq62ZLde2xvbaXPJc37mr53bPfW5nbse322N3055qH2K/wb7X/oODo4PIoc1h0tHAMdGx1vEGi8YKY21mnXdCO3k5rXY65vTW2cFZ7HzY+RcXpkuaS4vLo3nG8/jzGueNueq5clzrXaVuDLdEt71uUnddd457g/sDD30PnkeTx4SnqWeq50HPZ17WXiKvDq/XbGf2SvYpb8Tbz7vEe9CH4hPlU+1z31fPN9m31XfKz95vhd8pf7R/kP82/xsBWgHcgOaAqUDHwJWBfUGkoAVB1UEPgs2CRcE9IXBIYMj2kLvzDecL53eFgtCA0O2h98KMw5aFfR+OCQ8Lrwl/GGETURDRv4C6YMmClgWvIr0iyyLvRJlESaJ6oxWjE6Kbo1/HeMeUx0hjrWJXxl6K04gTxHXHY+Oj45vipxf6LNy5cDzBPqE44foi40V5iy4s1licvvj4EsUlnCVHEtGJMYktie85oZwGzvTSgKW1S6e4bO4u7hOeB28Hb5Lvyi/nTyS5JpUnPUp2Td6ePJninlKR8lTAFlQLnqf6p9alvk4LTduf9ik9Jr09A5eRmHFUSBGmCfsytTPzMoezzLOKs6TLnJftXDYlChI1ZUPZi7K7xTTZz9SAxESyXjKa45ZTk/MmNzr3SJ5ynjBvYLnZ8k3LJ/J9879egVrBXdFboFuwtmB0pefK+lXQqqWrelfrry5aPb7Gb82BtYS1aWt/KLQuLC98uS5mXU+RVtGaorH1futbixWKRcU3NrhsqNuI2ijYOLhp7qaqTR9LeCUXS61LK0rfb+ZuvviVzVeVX33akrRlsMyhbM9WzFbh1uvb3LcdKFcuzy8f2x6yvXMHY0fJjpc7l+y8UGFXUbeLsEuyS1oZXNldZVC1tep9dUr1SI1XTXutZu2m2te7ebuv7PHY01anVVda926vYO/Ner/6zgajhop9mH05+x42Rjf2f836urlJo6m06cN+4X7pgYgDfc2Ozc0tmi1lrXCrpHXyYMLBy994f9Pdxmyrb6e3lx4ChySHHn+b+O31w0GHe4+wjrR9Z/hdbQe1o6QT6lzeOdWV0iXtjusePhp4tLfHpafje8vv9x/TPVZzXOV42QnCiaITn07mn5w+lXXq6enk02O9S3rvnIk9c60vvG/wbNDZ8+d8z53p9+w/ed71/LELzheOXmRd7LrkcKlzwH6g4wf7HzoGHQY7hxyHui87Xe4Znjd84or7ldNXva+euxZw7dLI/JHh61HXb95IuCG9ybv56Fb6ree3c27P3FlzF3235J7SvYr7mvcbfjT9sV3qID0+6j068GDBgztj3LEnP2X/9H686CH5YcWEzkTzI9tHxyZ9Jy8/Xvh4/EnWk5mnxT8r/1z7zOTZd794/DIwFTs1/lz0/NOvm1+ov9j/0u5l73TY9P1XGa9mXpe8UX9z4C3rbf+7mHcTM7nvse8rP5h+6PkY9PHup4xPn34D94Tz+6TMXDkAACTtSURBVHic7X15dFXV2fez9xnuPXdK7iU3kBAhBMKQhEGCgLMg1qLFpdJWcOi7+Opq7Ve+1fb9anXhW/v6WupqreLbFt92aT8LONSWwmqrKNUutFCRQZB5NIEQEpKQO49n2Pv740mOlzvlBgJ66f2tu5Tcu4fn7POcZz/7mQ7hnEMJJRQP6GdNQAklDA4lli2hyFBi2RKKDCWWLaHIUGLZEooMJZYtochQYtkSigwlli2hyFBi2RKKDOJnTUAxIdVTSAgZkpafKxQF2eQyc9hygP4Lwv8RACAELmT5Oeecc0rP2ZEYY4SQzPvKGAOAAhtfSnDOkYw02hCfW7IzMWiWTeGJnKAp18iG4okgAHnWjQMwxgFAoDkbMQ6Mc0pI7iY5Bucc7xnnPBQKqapKKXU4HBaLJfXXtMaJRCIajTLGZFl2uVzmCIXc/tRmBXbJM0ImGGNprGm2TyaT0WjUMAxJklwuFzY7PxoGS2fhswyWZTkUIrA4z8di54Wsl8SBGwzET9mQB5OaL2FEVR43DALEKhCXRahQRJvUpwIxzgEKZVyc1Ofz/fnPf967d28gENA0jVJqt9vHjh27cOHCMWPGpPI0IWT37t1vvfVWW1tbLBZjjImiWF5ePmXKlDvvvNPj8ZgCe9++fb/+9a8VRcE/I5HIV77ylblz5xqG8fTTT3d0dFgsllgstnTp0kmTJgFAMBh86qmn8IFJu2WCIITD4UWLFt14442GYQiC8POf//zkyZMWiyW1pc1mq6urmzNnTm1tbSav7N27d8OGDSdPnsQnTRTFsrKypqamu+66q6KiAschhLS2tq5Y8awsWwAgHo9Pnjz5oYceSh0NCdi+fftLL73kcDgAIBIJz5s3b+HCL+NPfXeBMUrpSy+9tGPHDpvNBgCJRGLp0qUTJ07MfKLSMAhd1mBcoORPR3r+Z8cZi1U0OAdzQQgAAAXQdNbgVf57Xh3jQAm0heLffLuFQ+6N2XwEPt3G+7/kAAQoIVpSv2a08z+vG41i0uzKOFBCRAr+hPpOa+DdE6HdPYn2cDKQ5AmD9+0FhCgiqbAK48qt14+0Lah3z6hyAYDBuTDQQ4X36cyZM48//nh7e7uiKIIg4O2Jx+OnTp3auXPnj3/843Hjxpm76oYNG55//nlBECRJwnVXVTUcDh89enTbtm3/9V//VVlZicPGYrGWlha73Y53KBKJfPDBB3PnziWEnD59urW1VVGUSCQSi8WQGF3XW1tbE4mEeddNUEoDgUAwGDS/OX36dEtLi6IoSBg+FYZh7N69+4033vj6179+22234YXg7O+8884vfvELQogsy6lkHzt27MMPP3zyySerqqoYY4IgJJPJlpYWi8VKCIlGo16vN+vSRSKRlpYWp9NJCAkE/NOmTcvarLu7u6WlBTk7FouZF5sfg2BZZKpjfvXvR0Jgl7Ns+QRAZZ0JACAcOACJqOzt45GCBHMuUAJxHSSRAGGcmyMh+0Y045kP21/Y23s6oAEHECgIFCgBVMAIcA5xHU6FjFOByKZPQk9u7fpSnfMnc0Y1VdjxCcx3vZxTSl9//fX29vaKigpN0zjnsizrus45d7vdfr9//fr1Dz/8MOdcEITe3t5XXnnFZrNJksQY45yLoqhpmsVisdvtbW1ta9eu/fa3v433nlJqsVgsFgvqi2VlZa2trdFo1G63y7KMP6FE71taQux2uyiKhBDDMD5dckJkWcat3PzSHMFUXhOJhM1mI4Romvbiiy82NjaOHj0a+TUYDK5Zs8ZqtcqynEq2LMt2u72zs/P111//3ve+ZwpaHJkQout66qSpEATBbGa1WnM1kyQJmwGAYRj5heunHFFIo1RYRCJYRYtFFDI+skUULKJT/lQMUEIkiyBYRDFb+0I+OKZDogCQxq+HfNHrVu/7z/c6T0e5oEiiTaIyJcKnkrpfzgIViSALok0CUfzr0fC1qw688UmvQImRVymilGqadvToUZvNpuu6YRiyLD/zzDMzZsyIRqOcc4vF0tnZaW6Lx48fj0QikiRxziORyBe+8IVnn33W4XBomqbrus1mO3r0KO65OD7vBzJxb2/vsWPHAMAwDPOnvutlrLy8/KmnnlqxYsWSJUvi8TieiqLR6LRp055++unnn3/+xhtvBACUwTwFjLHHH3987ty5SLMsy6qqHjhwACcCgJaWllAoZJJ90003Pfvss2VlZSbZx48f1zTNlO78XGRdOp6BQprlZ7xP70uB7VKmAYPzPB927tz5Gxf4SRXoBucESFso8YXXDu/pUiWHhQhgMK4zzjikait9BAMwDgbnOuOcc0kRQwb96rpP9nSHKSEs92oCgK7rmqaZqqokSV6v1+l0mhsuY4a51olEAqUUbriVlZWVlZXm7owSTtf1rNOh0Nq/f3+uZaeUDh8+vLKyctgwD+fMJFJRFK/XW1VVhRphJhhjNTU1V111FU6N1KqqajaIx+MmhYwxr9dbWVmJGgs20DTNXITPAy66K4GQ9M/AXTJ7pbfg39zY0u7XJUXUjOxcRwlk3fY1xkWRxjX+75vaSN9shcKUWyQFKVea+m9ARaJA4YHPw4EDB3iGNc0ESl9N01K/xIcEf8o1uKqqqqqmkpc6RRovotqTNtrnh1/hYrMsBzAMzg3GDMYNxg3OjQG6EAKcc97XnukG4wbT+lcQj03vnwq8fSws2CTNyH6fCAGmMqbzrEutM04t0qYTkd1dEUogl6DNMTJhjOmfIuf1DMqiieawtra2rq4uWZaz8l/mQ5L/+/zTDQnZnwkunveLAIBEodohMKCk7/TGCcDZBNdyrBgB4BzKLMQuUcYBgAgUVBEqrKYiBUDg94d8hGUXjzgC52x6tdWXME4ENBBpuq4AQAnoGmxqC1053IHGjQKBe3F5eXlZWRnaXAvtORAEQQiFQocOHcLT21ANiyPLsoz/RhtZLi2iKHCxWBaZoK5cOfTgtLSfrn/14N7OBJVppslBoESPqT+8seab06pUg6PBlQNIlACAQAgnAMB3dMa4SLNKR0IADOOl20f/W9PwsKrftvbIlpMxQRaynbTIYV98UBclCEI8Hl+8ePHChQvx3lNK0X40qHHyYN++fTCkGzEhJBKJHDp0CMVnMBj0eDzTp08f2lkuJS5ujAElxCWn2xHFgVbKJlKHJECGYQTNtWFV74zqQEkmD1ICTOMTvJZ/a6rUGXPK4uxq55aWCCFCmqDlAEAgkDAgr18tE2goQLvM0AJHPnz4MKXUtCpcOCilP/rRjwKBAJ6oxo4du2TJEvQOFGhU+rzhoofFZDLWgJqjwYED6IynuLWgT94SiOtGQteBZNnuCQAwXm0XOBCGg+Sd7Lx3XzyB9U16wbKKEIInHlEUfT4fAIiiWLjRJz84593d3YqiGIbhcDgWLVo0ZcoUtMgO1RSXGBffYpDxOY8u5Nxf83UlPK4T0h8KczF2PpaCC7zraPnyer0ej0fX9YuxUxNCRo4ciTwai8WWL1++Zs2a4uVXKMZ4WQ6EQ3ZZzYEDIUHV4JxfJDUN91NRFAVBQCfWhYyG5liPxz1x4sRkMnkxWJYx9sQTTyxcuDAajQqCYLPZ1q5de+zYMTR9DPl0lwCXXbwsIVGNxXRDEQWAArSQwQCtpwcPHmxtbbVYLLquOxyOa6+99sJYjRNCpk6d+s477wwZoamjcy7LMobvoGPZMIzDhw/X19cXqaC9rFiWcwBCIqoR15kdQ7eGVGwxxiwWy7vvvrtu3Tq32x2NRmtrR19zzTUXMiYhVFXVhoYGp9NpGMbFELSc80QikfpNkcpXRPEpBnluKQcgAvHHjE8CiYsnQRRFKSsrc7lcZWVlGIV0ISCEqKrq9XpHjx6dTCYv0ine9CTjn7niVIoCxceynAPwnNGMAgGu8d/t6yEEOAAhIBCS6zPYcG8EY8zox5CIKxykqanp4p3A0mIbSix7SWEVqVWiuVIjDMaJVfzdHt/urjABiKjMSOrJpG6c+1GTupHUo7m8cLlRuJUAoxEKaYlcOnny5CG0bWUSU3jLwp9DbGwiRyvCORTQrFAUky5LCHAOLotU65K7gnEq0EyfFgegFBJJ+Ppbrdu+1jS5wnrtWKdVEY1zF4pS0JLGjOE2GIy6i7GnGKeX+Wvql4xxi0WmlKZJzRzBAxQAxo0b5/V6A4HAxWDcNDJSY2vS5sIo70LIRkPygJoMnlnRTThourOhmFgWAAzORUJuHeP6sDVCrAJkc5QyBoKF7m5PPPjW8VW3j186fUTeIc/JdEiDKIroiELbViKR2LRp04kTJzByhVLKeZ8jAAAwohn/lGV57959DoczFAqZMayyLCO7pzEERgLZbLbx48dv3rw51yNx3uCcW61WJAynTo3/x5/Q5iXL8oEDByoqKvx+v+mBkyQJrzcty00UxUAg8PHHH+NPhJBJkyal+e2w2ZkzZ/bs2YOHS4vFMnHixAu5nCJjWWSvJZO9P992JmFwlLuZMBiINnH1bn+ZpfUX88bojBMCFNJ9vCS37wpvoSRJDQ0Nb7zxhtfrxU3tueeek2UZI/xDodC8efNMA+f48ePLyspisZiiKHa7fefOnVu3blUURZIkQkg4HJ41axaGjaMYQ7NuqnF3ypQpmzdvRhbHn7LGbZldchmGU4dFTnW5XCgRsXsoFIL+EMRx48a53e5QKGSz2Ww228cff7xjxw6T7FAotGDBAlEUdV3HnAhBoOiJUBTl5MmTTz75JADgWv3qV7+qrKxMJZIx5nA4du3atX37dkJIMpmsqqpauXIlcnYqnVkvNiuKjmXBYHy0S/nRddU/2NguOeRc8Yc646JN+uXWLs0w/ufWcYxznlegZgJv9n333dfW1nbgwAH0HUiSZBgG3vIbbrjh/vvvR3HLGPN4PN/4xjd++ctfBgIBs3EymcRD1fz58xcuXGh69nVdD4VCnPNoNBqNxvBuNTQ0EEL8fj9qIJFIJDMkHDsaBgOAUCgUj2eJ7IlGo6FQSNM01LxVVa2uruacB4NBjOw5ePCgSbbL5XrooYeee+65TLIB4JZbbrnnnntMEavreigUtlg03p9XY8a/pz48mqYFg0FTLTbDIzG5w2wWj8eDwSC2icViueLf01BkLAsAAiUG4w/PHLm1I7J+X1C0S3qOxHOdccku/3r7WY3xF+ePw5yFwpkWb4bb7V6+fPmWLVv27Nlz9uzZRCIhSdKIESNmz549a9as1Mac8+uvv76uru4f//gHZqcYhqEoysiRI2fNmnXllVdiS7y1Xq93/vz5Vqs1mUyOGNGnuowcOfLuu+/2+Xx4ok8mkxUVFSYl+F+vt3L+/NswLieRSDQ2NkLGXnH11VePGTPGjLsVRdHj8SxcuPDMmTOYKCYIQjAYLC8vBwDO+ezZs1esWPH++++3tLQEg0Eku6qqatasWc3Nzalku93uL35xfmbUDjoprFYr/lldXX377bdbrdY0DccwjLKyMpPaqVOnyrKMvdIuNt99KVxtwjiVZ3d0/N+NbaJNzmQUgYChsplX2Lfd35AneXz6qv27O+JZgw9FSvSY+qvbRn17enVaWEwqMFsmpuvzXju87VRMVMRcXNs3ZlR94ErPqi/Vo2gYlG0rU/XM82uexvnH+WyRn2z4nIUpFp+Uhf7zikMS//Ll+jmvHjrYrQlWwcgra9fs7uUAq79Uz/ngmMcMs4KUfdDMAs84SBF+blEWs29mtYtUYWH+mmYDyswRSOuYNYkgbRAcPOuXF0h2GszGBTYb8GKzoihZFgAoIQbnlTbLhq9OuPnVw5/4dcGSk2s1xiW75eXdPpEee+m2eoNzCnlqK6Qjcynz2Gvw5JGnb/7vB7QEFXJfc53J8g97IWSfX7PzM3sVnyvBhECIzthol/L2ogmjygQjaeSpS6AxLtnl333ke2hji0CIwYY2YKaES4dilbIIkVKD8XHltncWT/zCa4dPBo2BZK38m+09Lovws5tqdcYEOnCBuTzeGjwmDxjFl7pdphVySyveZjZIGz+NDNOMmnUWs3GqGpM6VOZF5Zd2ad0zryKrGmCqUnmq1qU2ztUyE8XNsoDpYoyNd9v+tnhCgVz79ObOKof4vRk1OuPiQKtUyDoW0gbvSlp5orSOmQ3yTJF10qz3PpW3kO0K35Ez821MIs3RBlQDMk8PF+IJK0qWTauvIRCiMT7ebd+4aOKtvz9yMqgLFjHnaYxzwWb597+1T/DYbqvz5CrOhavs9/tffvnlrOMQQhKJxPz58xsbG9evX9/S0pJWto30hW977r//frypgUBg06ZNR44c8fv9WJKosrJy9uzZV199NWYN7N+//+2338ZqHehsu+OOO+rr61VVXbNmTSQSEQRB1/VFixaNGDHi5Zdf7unpMd1p999/f3l5OR6ktmzZsmvXrq6uLiw7V15e3tjYOG/ePLvdDgBnz5599dVXkSRVVUeOHJlqec1chJaWls2bN584cSISiaAHYdSoUXPmzMFiZJTSkydP/ulPf0rLZbdYLMOGDaurq8PYibSR//CHP7S3t5tOxHg8Nn1689y5cwesIQdFyrIZ2zmRCADwCR77lq81zHnl0HFfztMY58AJJ6L4vza07l3i8NrkPEnhsVjs3XffhWzndPQhTZkypbGxcefOnR9++KHD4TB3ZOiPKhw1atTixYslSTpy5MhPf/rTrq4uzGjAnf3AgQN/+9vfFixY8NBDDwFAe3v7m2++WVZWZppO3W53fX29pmnvv/9+T0+PLMvJZHL+/PkjRoz44IMPjh8/jkldlNK77rqrvLw8Go0+88wz27ZtM5MmOOeGYWzevPnvf//7D3/4Q6/XGwqFNm7ciE7/WCzW1NR0zz33ZF47cs+bb77529/+VlVVJBsADMP4+OOPN2zYsHTp0ptvvhkAzp49u2HDBrPAjKm3oLG2rq7u+9//fnV1tbl0qqq+9dZbnZ2dJstGIuFwODx37txC7n7xsazBeHdMZQAYAtVfJZETAJ2DVRDevmfCF18/ctyvU0nImjjOOIgi7fJrP/mw/bmb6xjL52DAaodYF83Uz7DqhuljtNlsZWUum82OSp6qqiifsKMgCJqmPf/882fPnvV4PIwxrHRksViQCd58882ZM2fOmDEDS2RirDem6xw+fAQARFF0OBzJZFKSJItFxl52u93lclmt1tQCbH/84x+3bNkyfPhwdC8nk0lRFLFm49GjR9euXfutb32LUup0OpFlRVFE0Zu+RIxRSltbW1988UVZlrGoaDKZxIvFHWD16tXNzc3l5eWCIJSVleHmgJePqZG4VocOHVq1atWyZctMNd3n86lq0uPx4BJhRT10YeTSi1JRTCyLvit/Qp25eq8/SWlGji0hRFf1bV9reHvRxObf7g/pjAgkRxACJxZx9UH/f1yjVihypuMDF3fYsGGPPPIIpfTw4cPr1q3DCJJoNNrc3Dx//nxVVceNG4ekYewsekeXLFlSW1uLIgeDmD755JNTp045HA5d19FltXDhwueeey4ej2Pk1MGDB2fMmIGRuGaQniAIp061dXd3V1RUoKsT02DwyUltDP3a4UcffeRyuTCWVxTFxx57bMOGDR999JHdbrfb7VimDv20feuQI+QXp9i5c6emaXa73TCMWCz2wAMPKIrywgsvKIoiy3IoFDxx4sS0adNMSpCt6+vrnU7n7t270YfncrlaWlri8ThuCIIgnDlzJhqN2mx2dN4iK/f2+vx+P2ar59eMi8/IxQCCGkRVCKsQOfcTViGuQXdcH1umvHh7Ldf1XJfHAahI/CH9/VNB6C8Cngmr1Tpr1qyrrrqqoaFB1zXc8nRdHzFixIwZM6655hozCuRT8hibNGlSY2Pj5MmTJ0+ejFFLwWDQMHToT6mtrq6+8sorrVYr6y9Hl9W9LopiOBw6ePBgIQmxWAEEBSH05/w0NzfX1NSYReAKdOKbwJq1OLUgCHPmzGlubjavlTGWWo4OaUgkElOmTHnssceqqqpUNWnOi5k8OFRHR4eun5MyRCmNRqM9PT1mm3xXOqhr+JxAJEAIlpE954PfyAJhHL48wXvvlGFGIqexlgAQDrvOxCFvUqOu64yxRCJhLjGyHe7vWevEYAlE7Ihcknp7zO4D3huMbNy3b2/+Zrm7c8MwhqpoIWMsGAxiFA5JQVozQkg8HuecO50OwzjHMGf+u729PdOAoGna6dOn4XJlWZ73AwCUAAf4j6urLRbCclTvAgBO4Ew0CXmjvFHnS1tiNCRlfp/2a664ZmxAzkUW8jiTZcuRI0fD4dD51Y8xicwzSyHAxwzDz1FFQWRlrzy2alyNzs5OQejbN8wROOfIsgNf1Pldw+cfBuOThtlnVNu5auSJg9EuQsGDAmtzqP0wI/3SwDmXJLmrq6u19YTFIptlZQcFXdeTyaQ516D6pj4noihu2LBB13Wn06kois1ms9nshZyWUoHStLu7WxBE1FmdTidaYwRB6OjogAJMtsV0/BoUsMDRaJfwT8ZJrlDwiwNFUfLnjSAfL1u2DA/7hmFgHGA2fwGJxZL79+8XxUGnKqBmeeedd95www2iKGIUdoF9UR5Pnjx57dq1AMAYs9vt27dvr6ysfOGFF81MYEVRspKdBz6fD6U1Y0yWpbq6Mbt37xYEQRTFrq6uQowGly3LQl8Zr0saSsA5F0VxzZo1+AaiWCx28803z5w5M9M/CQD19fVp3bMa80VR2LdvXzKZxBSxQYExNmLECDMet3CgVWH69Ok33XTTpk2b3G431hn561//qijKAw88kHohmTRDbj7u6urCxA1d10VRqq2t3bVrF0pZn8+HhRnzGw0uY5blBRcBG0pQSnft2oU2Jr/fV19fP3PmzKy3Vtf11GCDzKwptDhLktzW1gbnW1suNW4hc5Y8wC5Lly71+Xz79u1D25nT6XzttdeSyeSDDz6IR880ocg5x5DtYDCUVV52dHRommaz2VDK1tRcgRMJghCJRLq7uwdk2ctWl/0M4XK5hg0b5na7Kyq8ZqD+oGAYhsfjURQrcpthGENS4mBQI6Cd32KxPProoxMmTMC8S8aY2+1et27d6tWr0cmS2oUxpijKzp07ly1bllqO3Ew9AID29vbUwYcPH45KAqq5nZ2dMJDR4DKWsp8B0Pfz3e9+d9y4cbjuWC87q7xJFXhpAdeqqtbUjFRVdc+ePYpi55xdyGF/sIekNKpcLtdPfvKTJ554Yt++feiZc7vdr7/+enl5+R133JHannMuSRLuCZjqg5EVzc3N6KUTBKGjoxN53TAMp9NVV1dntSqRSAS7I0PnR/GxbN9mnyPvA1LUV9PmdSnBOR82bBimMeVHS0sLShfD0F2usuHDh/dLF8KYoShKff34nTs/stvJeRcJp5T29PRghjdqBbW1tQVeBSGkt7c3FAqhSW7JkiW/+c1vjh07ZrPZUENYtWpVY2Pj2LFjM4WiqqroOxAEYfz48Q888AA6ZnVd7+7uQg2HUppMJjHz1qQW7Vz5n8/iY1kOGe9J6geGHMgiMTjXGcgCSDRvQCwHKWflz/MHvosG2RH9rmkNUHotX768t7dXUSx+f+DLX/7Kgw8+mHJKI4ZhTJky5ULKcDDGrFbr6tWr169f73a7Y7FYTU3NypUrC+wrCMK6devWr1+P0TZ1dXXLli1btmyZ3+9HL3QkEtm0aVMqy+L+MHbs2AkTJkiS5Ha7a2pqmpqa8O0PlFK/3x8I+PvNBfKpU6eefvppq9WKKgcaDQYM5io+lhUJFQWa9WW6hBBgfHtHZHaVyyYRg7NDPhVozpd7AUClTR5yCgf0ESBEUZQkSRQlSZIzIlxJMqmOHTu2oqIiFApdCONSSiVJ6p9rcLfb9ImIophIJLxe7913371y5UqMEEIRDilCEc1qU6dOve+++1LHMYnv7u6ORmNm5q0ZNYGqAhYxx/i1PCewomJZAgBQZhGHKWJPRCUCTbuPjHEiiT94r31Le6TaLm3tjOzoSJBsqbzQ97oE3lChXALCs4L3Ietpg+i6JknS+PHj//nPf15g/RiegvMeBA9Mo0ePliQxNXUisxlGciEjms8tbiCnT59GcwGG0VRVjXA47D6f3+/vK6EQiUR6enrys2wxWQxI36uf6RSvhRhZ6mhgTcQkF/64P/DfH3ZvP50gYjpbm0Mxxq2KcH2NE/qL0AwV+LnI2mDAamrYEd9rcN5nr0IiGQoHnpkGbIYHPjNgN5V40yWL0UXf+tb/fuqpn951192xWAxbqqqKPrA8KCaWhX6lc9HEijzeSwIgKKJok6gk5LpfAiU8Ydxe56otUwzOz69qZy70FwISsp7WeX/1tfyMiI6DxsZGjP07DzJM1hkSrkUlQVGUVM2/8JFR8+no6MAFQXOYx+MBALRzmRHGAwbHFBnLipQwzheM88yqtetxXRKyRZOA+Urb7JctUGIYXLHCj6+/Ik+JkPMDIcTv9/v9/rNnz/p8PnzDjCAIyIKmxnby5EnzJZ1Zg0goJQAwcuTImpqatNd55pna1J5RYrW2tvp8PuRaQgZRZgWBzx7SHAwGd+zYsWPHDnwNBc6CDttChkXrMiZloLTGmtIAUFFRgS81R/14QKNBUemyAADAAURKV9029saXD3RFdFERGeO5i8WcA0qAUqKrjDD9d3fVTRxmY7xQrSDPocr8knMuy/LKlStpf+Ett9u9YsWK2tpam82O6qmiKPv373/kkUegvzSxqqqjR4/OOoUgCA0NDfi2cqwcaE6d1hh9aTU1NadPn8bzja7rjz76KADY7Xb0HqPnlp37Dt6szIFfjh8/HifFxIrly5ejoEWJq2natGnTMinJHBAfmEAgEAwGsTodljBDjne73U6nMxwOE0IkSeru7uZ530k2aClLCRFpvk+eYgIIIW93kQ5QgEgghHE+wWN7995JzVUWPZJkOiME+mYnBF+4bH4EYg5LmM71iDrKSf9yT/1XJ1YarFCVAJNSTKS9+1jTtNRfw+Ew3p5AIBCJRFRVLS8vX7x4USwWi0Qi5uuYGWOxWKy3t/eWW27BV9NjVH8ymUgmk6raN8XUqVMxRLp/eJX3v+07dVL88t577/V4PD6fT1VVM348mUz29vZeccUVGBuAqRNmx6zhXRhUPnPmzOuvv76npwcjgK1WqyzLuq7H4/Genp65c+ded911KWT3ITMqDWnr7Ozs7e3VNE1V1Wg0iloBBtzY7fZIJJJMJg3D6OjoCAQCkFt4D1rKxjRDj6o6AKSWHCQAHHQKkGT++ABmI39c06Mq6LTP1o9MwwEI6JRAVI0PVFwbS8U0VTj++bWm3+zu+n97zu45m9A1DoB8Cp/mcnEAxvC1tiBCvdtyb0Pl0hkjKhQ5T82vTNjt9rq6OpvNTgiJRCJerxdSNq/hw4fX1dUpii11lVHyYWoUACxYsKCmpmbjxo1tbW3RaBQNk9XV1XPnzp0zZw52dDqddXV1TqczGo1WVVXhOBMmTJg2bVoikUDBo2k6OpZGjhyJPk+Umvjl2LFjf/azn/3lL385dOhQMBjUNE0QBI/H09zcvGDBAnzjrtVqra2tRWU6kUiYiYSZEEXx4Ycfbmxs3Lp1a09PD77mCb2s11577a233kr6U9/q6uowyygcDiMvZiISiVxxxRWYExYOhydMmAD9YQ8NDQ26rmOGma7r4XDY7XbnomoQ+g3jnBKyszP43omAlC0TkBAwDF7tlBc35Awd4sBfOXCmK6IJ2bKyKCGaZswb4542wjXg25CRHgBg3NjRGd16Ory3J3YipPXGjZim6QwEShSJuq3iaKfcVKHMrrbPrHbhy5Vy5YLnQf5YDchqKM7WHWWMYRhWq9V8O8hAgxdahc5sxjkPh8OYGYuRjYUPkhXRaBQdWhgsex4j8LwV6VIpz9MMBsWy0JfIWsDC5b57F7JqWSdijGeoIlw1GDK0JKS7v3TGhaG1aRWMtNow0G8Ou5A6FJnIOqapv57HgJl9LwbZhWPQp0jGBzjrEID86qyRxxkFAHhIGszicgDGOedYlTu9L+eAxkn86TNh1nPpOUd5uAQTDcksl4zsATFoli0KpF7SZ86jJQwtis/IVQhKbHoZo8hcCSWUUGLZEooMJZYtochQYtkSigyDPn5dhvaFAo5rQ2tOLiE/8q/25WnkKuEyxuCkLO9zLAEh5LKRtxy4SImU25eDIfTnEbxXwnkD4x6z/1TgEOiDTers5t8fPhXSRJEUGO/3OYdISTKu/Z+Z3h/MHpUZKIOpc6+88sp7771nFuku4aKCUhqPx7/zne80NTVlTV0ctC7bGtI6AxqIl4uUpQRiam8iX+HVYDDY0dGB5VIuGV3/ssDQXrNQbiYGzbJ2icgypZeRlNV1ahHyWU4wLhuLUF8ywv5lgaG6eWJuBqvLwtm4rkbVy0bKqpRAVIuo+aRsLBbz+/1YXv2SEfYvC6zonRZEn4qCi4phawKPzBzuT7C09xSYUdq5vjE1RJ6jQeaXmV2gv7hGgZMWMq9AiKoa88Y4IFueLZ4Arr766mHDhqW9I6mEiwTS/2onyBEyVjoFl1BkGLQuO2C0azGCknwJZ0NbDaCEQpAnab4kZUsoMpRiDEooMpRYtoQiQ4llSygylFi2hCJDiWVLKDKUWLaEIkOJZUsoMvx//c19hxcHq4UAAAAASUVORK5CYII=";

function detectDisciplinaForPdf(rows: any[], title = "") {
  const disciplines = Array.from(
    new Set(
      (rows || [])
        .map((r: any) => cleanPdfText(r?.disciplina))
        .filter(Boolean)
    )
  );

  if (disciplines.length === 1) return disciplines[0];
  if (disciplines.length > 1) return "Discipline multiple";

  const match = cleanPdfText(title).match(/Disciplina:\s*(.+)$/i);
  return match ? match[1].trim() : "";
}

function headerValue(value: any) {
  const cleaned = cleanPdfText(value);
  return cleaned || "____________________________";
}

function getImageFormatFromDataUrl(dataUrl = "") {
  const lower = String(dataUrl || "").toLowerCase();

  if (lower.startsWith("data:image/jpeg") || lower.startsWith("data:image/jpg")) return "JPEG";
  if (lower.startsWith("data:image/webp")) return "WEBP";
  return "PNG";
}

function isAllowedSignatureImage(file?: File) {
  if (!file) return false;
  return file.type === "image/png" || file.type === "image/jpeg" || file.type === "image/jpg";
}

const SIGNATURE_DATABASE = [{"name": "Arch. Sergio Raejntroph", "discipline": "Coordinamento", "role": "Responsabile Tecnico", "signature": "data:image/jpeg;base64,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"}, {"name": "Arch. Veronica Laino", "discipline": "Coordinamento", "role": "Coordinamento", "signature": "data:image/png;base64,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"}, {"name": "Arch. Veronica Laino", "discipline": "Progetto Architettonico", "role": "Ispettore Tecnico", "signature": "data:image/png;base64,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"}, {"name": "Arch. Arianna Brunetti", "discipline": "Progetto Architettonico", "role": "Ispettore Tecnico", "signature": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAK8AAAAmBAMAAABAA9ZWAAAAMFBMVEUAAAAREREzMzMiIiJVVVVERERmZmZ3d3eIiIiZmZmqqqq7u7vd3d3MzMzu7u7///9re723AAAEGklEQVRIx+2Wb2gbdRjHn6ZNL03S7jbXuqFbUuYEhyOHINUXW1PRraDQ2wpCkZlbJ6gIS3UvhO1FDvZmr0x0G4ov1ii+EiEd85XK7vpGh4iJ6PZGIRmiDim7tUkuyV3vvv4ul7TpaMsxt3f7vbg/v9/x+T2/53m+z3OEBzPu0EPw/wF//8qDARs85TZZvn7P4OzTeniTZVnrBP/hnbvsg9W/yXp0DXibdzAz1wptum8H2OryDp5Lwti9SQQCneB6V94zWFBQTW28XO3vBJfPpj2D+Tw+0zZY+ysHdQAVZQV883rCM3grMLLRmpjATAyZ9Ap4rhLzyrX9qG70cYMSENIQfsq1wef1iGeLByGt74lGsSykLdLqveaBNvg5I+zd4uoGWVwY/bCgVPwohGEWXfByvx3wnhVUXH9HfmLncajcj2LPeNvH9VFMes8KemddV9Tok8c5zBDR3uE2eCmJrOdEFg7x3ZfWmV88JZ8Lg/d/7YPcBl/OQ02yu+YF/CzsL9zy9uaa+YzsK03UKLbYx+LbAj/BKsAAC2u/F/BRR16c48BuGC1d3WKSGBKDBWVeTon9y1wLvMyARtC+IEx4AYvOJb51FmXS5sL423mVWCL08rFJ+9HSN3siZrAFrjDZmT6R/LBfXcP4nXlJwW2gMwBxx2Ey9eAXQZt+DWoO0vP7RqAT3eaMSPbzy0kz5IJvyS+/f0KgnTfD+LMb1gfsCGNMtFdhCxqEHMY09uXCz2sszpC0a+wIDo/jTrIcFJVDzP6AES7l5ICaNsMuuETUNXhGUgqsgkxqOttufmjsIJOM6dMsFiZB0X0f8fR6Cyw184isKPVh5DD0iKpIylEUaLQW2IFoQk7bLR9XuShzSiH9LruchdqHGz2we8spLO3TGlwSfLFOvm/PDa76x6kKebOriOHHYAVVRfgyAIlyOqWxF/E0pl1wNmUdP4nF2Avs8T3MccvEjM7E81h8UdN9Dth0ZuZdwdmXXItnGQMzozDpyn6iWIN6UW7KN1p00paBbdaWrCPFyhYGlp9EKVjj2FHKxHLbhwopDGz4v2MCcMGW03r+JUphRkMmieozTGvHEOcTWGom36l2oW+aUA/qW2QFhRADN0J+F1wKwCANPBpsAifaQuh+4yVKSTGox27QQWSUi29vx5VdjZWa0QZfcJv2tXA5ZIsRpwMMs9wQpLTzaJKmUy5LJ+3z7QZqX3xk6ipKo2zTrq9o2Dm+LAzdne0M/Jvbe7eFbZG2syCEoLLDHdCDiPfhV5oSpqmnQeRfW0yMWWAhj3+mHAcZU1gH3BrZ3bA+hcxz0hnpqenTsKM8J+7x/0D77Y8VXDvtvdveBV4caBoyrtX5Ha5Nb2n1wbyju3sZq+DaSnNawH0YHf9uGu7nePhHvwr+Dw7fb/I3Uq6sAAAAAElFTkSuQmCC"}, {"name": "Ing. Salvatore Grimaldi", "discipline": "Progetto Strutturale", "role": "Ispettore Tecnico", "signature": "data:image/png;base64,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"}, {"name": "Ing. Bruno Gabrielli", "discipline": "Geotecnica", "role": "Ispettore Tecnico", "signature": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAPEAAAA9CAMAAABLP3/0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAALTUExURf////7+/v39/fz8/Pv7++/v79zc3M/Pz/Pz8/n5+fj4+Pr6+tPT0/Dw8KqqqoKCgrCwsNbW1vf39/b29u7u7jk5OaioqPX19Z6enp2dnZSUlF1dXcnJyREREaurq4SEhB0dHZWVlTg4OIWFhW1tbRAQEPT09NfX16amplxcXOTk5GZmZiIiIrW1tampqVpaWt/f36KionJycr29vfLy8urq6uzs7Ly8vMbGxuvr6+jo6OLi4peXl19fX87Ozt3d3Y+Pj+Pj42dnZ8jIyKOjo5OTk4qKihMTE62trXZ2duDg4Obm5uXl5bm5uZGRkVZWVrKysoaGhtjY2DIyMnNzc9vb28TExMvLyxUVFW5ubsLCwk1NTWJiYp+fn2pqalVVVVdXVxsbGysrK9nZ2VtbW0ZGRqWlpW9vb9XV1ZmZmXl5eRISEk9PT9LS0pycnHFxcczMzIyMjHV1dSwsLJKSknt7e1hYWFRUVNTU1M3Nza6urlNTU8fHx/Hx8X19faenp+np6YCAgIODg4mJiX5+fsHBwbe3t3h4eAQEBIiIiIGBgU5OTru7u3p6eqCgoOfn55iYmGtra7Gxsb6+vmBgYA8PD8rKyr+/v7a2toeHh0REREdHR5aWljo6Otra2mhoaNHR0a+vr3x8fD4+PjExMZubm46OjktLS+3t7ZqamsDAwKysrKGhoSYmJnBwcEpKSri4uGVlZWFhYTU1NY2NjXd3dzMzMwAAALq6uuHh4UlJScPDw0hISCoqKrS0tDY2Njw8PN7e3kJCQmRkZFFRUT8/PyAgIMXFxVBQUGNjY0xMTBgYGLOzsykpKWxsbFJSUtDQ0IuLi5CQkKSkpBQUFAEBAQcHBy0tLUNDQyUlJUFBQRoaGigoKFlZWQgICEBAQF5eXi8vL3R0dCEhIX9/fxcXFzc3NyQkJD09PUVFRWlpaSMjIx8fHxYWFicnJzQ0NAAAACUp4JcAAADxdFJOU////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////wCpCmekAAAACXBIWXMAABcRAAAXEQHKJvM/AAAH20lEQVRoQ+1aTZLrOAjWCVixzQnkO3AUrdlxDao4ge6Bc70pJP/FHTtJT2e63+v5qtptSwiBkABJSel/fCOQZF/0NwNFEBx+i86ICUdIKTGnhPvavxPIZXrZ1/ytkGuY+DdBL/R7zBvQUfdFfzfQr7Qv+3uBCZE44y+a05hSHqB76d+gNiZMUEr8+xw+2+5bwW6Rg+yLz4EpEaKZIP5hPgATcXk9ucSEpIpFkEXxzzF1k1MvkX28KjKBU7MuipU/ye9hEi1zvvWk3H3th2X7O415T7KHmL0+j74Yi3ZyeSjvDpiS1LLZdOTSND+BlMsg5yS99pzkOTxgRKXOFPF//jtCI0POlWa6+PRzhcPLlUs5UjkaLxPmvPsz3Cq6cGqz0JALo+ZsolJKXkR5sjP0gRd9UyL34HDeGEu9m7rfKfoKxDyEpFUSIiYSZAVkVxF1VaeEpmrPdY5JPG92WpjkGlvN09ZNgAMbY0JjV3D1XHkq2RM9hckIYpaJk0gybr71hoZz9BF2N6FYm/NUOEBU+LXvO2YiH0+90syQ7sXtqNQs0b2JAehYF30PhbiHttyk6MiJiAyPRLK88aEIxHbaT6vzOm5oMFGOAThrNuMjTZTwZbNxQ2O9NzIH6G6lzVytgGhkc+ubibKOoWSea9qTdKPybJzOtvFFTAIqWxYIsCzQjyo9AiYYdgOWa/t6yGsRG9GKE8ndrOBWmySjh4k3lLPJN6PQXhDZNDSTMtpcGA/StlZW3Ov2DO67Njy8wCIMoR6WaF9T4S3NFnbNH+o/uI0wrhtQ4hp++TI5l/ZEqhGZ2lc8kOikuz2afXi34vClnBcT5rovXOVbnvNY5OGG+11hsbsWxAQ1oY/h2ZdKyvOabmxhzHuLPULu6cCKKbg/yQQTxs5ver/RYc+hSThGstSaIcrqyHdPi4CNKZWcJLdT3qmCRH3Ll4YeXV4A1dlthQhmADccT9FERyQDs8VhHQKTmUEdjYwlpiINZTNKmGgZaaw9DeURkbMtiRFa7W51EtEsO9vHsT1AJwMOToISIvuleBbkvrN5gY+rSihECMoGYGDAbMBmbGgBitN4Hbxm4cokCbi6rWtQRFQmI3NzpmKFE1zHJbsSD4WnFjFLqjn0OfM0MGlJxnEfwuAibQSlMM3uf5f7Lsxve+mkxiCUSwZVVvDs7M6QmQAAuKU/ellWMV81iluyjIk8q5TmkjH8OSXUypRo2mbFhFbV2WdHAbEBeXPj00CtdSegzClfVSGJtESuU1MOTkct75XfK7tFpHYtOWwfNLoxk3cd0BwgT12iD5ZAPXJQqdpXNPA0ofsIhQPKEVnXfKlXPRREKiRjUNsdpFjexbwFK8cn2G8RjijiXrQxdxDGxW2iISwLkkYHhPjCZBcPOUhzXFBNBDES1mL1mD/GugegI1cX3jLQBn1WbxJv81wUn9U/6R5pgBa/k1UVHAmXHRAyLJmeZg11Oy+rQEl82t3PzNFz9wDDMMWJk24XTJT64fanh3e9jl5KzZwn1W+V7IllvCKsY9a2rXcRLaSO06oN0yIYTJEVS7EhN45I2WUJF5js6l5ss8eKBW08THGNhyVlXbH/vkXM6g1FvCKiFs8KEgcnJFw8zpYdtNV3Y0ei1hXXkqcUGFjdFeb8a6Ze+MMFJGisn9lgaVvJcO9qCWpRjihsmPSqXX0AGsebwBeFktc9URkjX1rtfK5tSxBjmLe2w4hxRDwsbhHRL4budsnziEiNvS0iQmEYh+KlFChVCE3LCGZE1AL1VlgaObERWS2llKpKVrPlWvoEJYlGnTJf6piLZ0DkOk+xELT1GVSNMDYZpYxrhLub2O/gdW7dOFrNVWKjp3lqHP/s4sViCyMWJinCkmRU1hy9IxFJ5IQ9N0Aiv4zjWJp4a/KQvFJqXBHRByBDEkzOd9LiYEIYaWY7+5htiGDrnJxaUR0GJxSK/NqawT/ym9DmpMfSmr6Fc46D87YeZ56hkQ/NrWAi0BGsHyx0fe5i0nU7xTDBhVPzu60cIKaxnYjXEIG6KpmIkBhkSAuHhQQlD14Ka4EYo1PgbOOgk0iBeiH69o6EtM5JwFr4gPMtGjOvIjANYhN0rjhDxCHVMtacw6/fyrDFQ04NQeMtJYgZW1oq21whNMeyUN1a6zMIh5EJV6//Cg61fBVdD46gEAMZAaHxhWUwZ7Lt4xNoEsMAeHC2csx3qlkJjkkbZjEPycKe1SPote1EgGjOPGaKO68voembsFaMqDpzeZ7bI8pH9bfAJJfYTpeeCRCD0ZTMH/Dq8r8MjCxiSXUWBi/of0pyZtY9JA71UIuxahJo4fOo8bb8Fc3Da12I+lHLnUZ3iiYc19zDoZlWRNrarmYosqsP2ea/xNR3eInIp+O841yc9yNWld0c+bxDoriQqAXXje33AiXf959fhL7AvLYDlu9HpEX28vbydWhX+N3dPAmIbPItWPwIVVOitp6/V+m+E9FXr65fQvMUIx9eQf3XaAn0Np/8ajS+dX9l8o0IjV8+z38Rwrn8GIXbfnOIJPNtQDKdd6xv7OYF0LBc2n0tZg0lKzy3j/uPAOvRztej3R4tv2L6EUCM44C3zbf40dV6A/FDIPI2iWJel2s/hX3TkH4C8XuXfdlXARHLGMfqP0fdAL3DxpOK1H5xtiv8ZkxSvEUYTP06/PfgJ63ffwANvWX+/CJobgAAAABJRU5ErkJggg=="}, {"name": "Ing. Carlo Renda", "discipline": "Impianti meccanici e relativa documentazione economica", "role": "Ispettore Tecnico", "signature": "data:image/png;base64,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"}, {"name": "Ing. Gianluca Biaggioli", "discipline": "Impianti elettrici e relativa documentazione economica", "role": "Ispettore Tecnico", "signature": "data:image/png;base64,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"}, {"name": "Ing. Marta Dominijanni", "discipline": "BIM", "role": "Ispettore Tecnico", "signature": "data:image/jpeg;base64,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"}, {"name": "P.I. Mauro Garofalo", "discipline": "Documentazione economica opere civili", "role": "Ispettore Tecnico", "signature": "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAeAB4AAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCACvAWgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9UKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACkY4xS0jUnsAwzAOFwT2z2qSvnvSfEWo2/7b2v6CtxI+k3PhO2unt2ZyqSxzYVwM7QSHYHjPvX0JXVWw7w7im/iSf3mFKqqnN5Nr7gooormNwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKyPFniC38K+G9T1i6/49rC1lupAGCkqiFiASQMnGOvemk5NJbibsrmvRXy9+xJ8YPGPxktfG2s+JryS6tF1CJbGIxKkVsrK7NEhVRuwDH97nketfUNdWLwtTBVpYer8Ud/uuYYevHE0lVhswooorkOgKKKKACiiigAooooAKKKKACiiigBrMV7Zrn9N+IGg6x4u1Twva6hFJr2lxxy3dicrJGjgFW5GGHI5GcZGetdCetfLPx4kg+En7Rnwz+IwxbWGrSSeHtZm5VAjgeU7t04OTyMkR9eK7cLQjiJum3Z2bXm0r2+exzYirKjFSS0vr6f8A+p6KghkLJ64OOuanrhXmdIUUUUwCiiigAooooAKRqWkbpQB80abdeZ/wAFAtXjikyq+DUWVVfI3CVCAVB64bv6ivpivk3wfNDP/wAFE/Gght1t2h8OpHIy5PnOVtWLtzwdpReOye9fWVezmcXGVFP/AJ9w/I87BO6qP+9IKKKK8Y9EKKKKACiiigAooooAKKKKACiiigAoprsFXmsv/hKtI/4SJdAOo2w1toDdLp5lXzzCG2mTZnO3JA3YxzRq9hNpbmtRRRQMK8A/bi8eJ4J/Z/1qBdhu9cZdJgWRSwPmZaTjB/5Zo/069q9+z1r44/4KGJe+LP8AhWfg3TJA93q2qybbTAJZ8JFG54JAHmPzjHXg8V7eS0YVsxoxqfCnd+i1/Q8/MJunhZuO7Vl89P1Oy/YE8L/2H8C49TK7Drmo3F+ihtwWMEQqB3H+q7kn3r6WrF8F+FbPwR4T0fQNPUrZaZaR2kIY5O1FAyfc4zW1XDjsR9cxVTEfzNv5dPwN8LR9hQhS7IKKKK4jpCiiigAooooAKKKKACiiigAooooAK8n/AGnvhu/xM+B/ifR7OJptSjg+2WKpnd58R3qByOWwV/4F0r1iorgZj9eQcE4z7VtRrSw9WNaG8Wn9xnUgqkHCXU8x/Zl+JDfFX4LeHNdnIN/5JtbxckkTxHYxJPOWwG/4FXqdfKX7I9wvgv4qfGD4cSsqNZ6w2q2UWxU3QyAAlQGPGDDwBwCM4JxX1bXbmVGNDFzjD4X7y9Jar8zmwdR1KKct1o/loFFFFeYdoUUUUAFFFFABSNS02Rgq5NAHyT4D/wCUiXxA/wCwGn/omzr64r5d+GOmpeft2fFi+8zBtdJs0VNinO+K3yd2Mj7g4BHvnjH1FXuZvJSqUl2p0/8A0lHnYFWhP/FL82FFFFeGeiFFFFABRRRQAUUUUAFFFFABRRRQA2QZHNfPPwQvrLxN+0h8a9RubIHXNNvLTTIbpm3CO0ERARM/d3NGWbHBJHFfQzV8o+G9QHwm/bl8S6RdXfk6V4506PULcT7gGu0/gViMEnbMB9VXqa9fAQ9pTxEV8XJdfKUW/wADz8VLknSk9ubX5p2/E+riwXqQPrTq8e+Mnxe1X4Z+Ofh5biwtp/DPiDUTpV7dOWM0Mz4EO0DjGSeT6Y4617APu88V50qU4wjUe0tvlodkakZScVuhkjBQ3NfG3jLXovip+3p4K0KKIXNn4PhklmZmCjzxG0rMCOu1jAMeob3r2/8AaP8AjPH8H/BrzWKLe+KtScWWjaao3yT3D4AbYCCVXIJx1JA718y/sYfDe48P/tNeNxf3ser32g2BhvL5SVP22ZkMwx/HgrOuemVz3FfR5Th1Sw2Ix03tCSj5t2i38r2+fkzyMZV9pXpYePdN/LVL8Ln3pGOn071JTY1wo+lOr5VHthRRRTAKKKKACiiigAooooAKKKKACiiigApsi7lwadSMMrSewHyx4+jPw7/bi8B60j+XY+LNLk0qctKVDSxg7QcjHXyMLnJOfavqivl79u+zudI8GeEfG1kWN34V1+3uwA5GUbg44ODuVPm7Anr0r6U0fVoNc0uz1C1dZLW6hSeKRTkMjKGUj6givaxn7zC4ev5OL/7dd1/5K19x52H/AHdarS80/v3/ABLtFFFeMeiFFFFABRRRQAU2QblIPSnUyQkLwM0AfJ37O15Peftj/HRp5WmZDHErOckIrhVX6AAAfSvrMsBjJxmvkH9mO6ivf2vPjrPDIssTS4DxsGBxNg8j3Br6x1q6ax0m7ukAZoIZJQrdGKqSAfyr3s5jbFRj/cp/+kRPMwEk6Ll/el+bLoYN0OaWvL/2cfilefGf4S6L4w1Cxg0+8vjMklvbOzRqY5XTILc4O0HB6V6hmvGrUp0KkqVRWlF2fqjvp1I1YKcHo9QooorI0CiiigAooooAKKKKACiiigBG5HTNeG/tTfAu6+LnhGDUdC/0fxp4flF9o9wrBWdlIYwlsgAMVUgnoyjsTXudNcBlIPSujD154WrGtT3X9fiY1aUa0HTnsz86Pjf+1BF8S/gPZ6Jr1lcaH8Q9G1u1mvtOkieES+UHzKhI+XkjKnlSeMjBr638d/tTfD34feEYdU1HX7OW/ms0uodIsrhZrmUsisqAKeM7hy2BjmuA/bm+EPh/XvhPrfjFdHiPinSY4nj1ON/LkEAlUSBuQHARmwDn2rV/Zy/Zq+HOi+B/CXiyHw5Dca7faTaXUl1fSPciORo1dmjSQlUJLHkDjjFfV1qmWVsDTrcso2lL3VbdqOl3tHto3ueDTji6eJlSTT91avsr9OrKXwb+HeufEvx5J8ZfiNYNaS7FHhrQrpsnSbcZJkdSAPMYHPTI5PHygYv7CNxL4k1X4ueKZkWcapr5Ed+EAEoG9yAeuP3inH+0O+a+l/F2qjw74M13UxF532GwuLrys7d+yNmxntnFeAf8E89Lhsv2fI72Pd5mo6nc3EmTlcjbGNvHTagri+syrYDEVJK13CCS2SV5WX3a93q9Tr9iqeJoxWvxN+uiPpuio5pUhid3dUVRksxwAPWuc0/4k+F9Y1T+zLHxLo15qOWX7Jb38Uk2R1GwNnIwc8V80oyldxR67ko2Te509FIGB71FdXCWlvLNI4jjjQuzHooAyTSK9Saiua+HvxA0X4neGYPEHh68+36TcO6RT+U8e4o5RvlcA9Qe3PWtjWNYsdB06e+1G9ttPs4V3SXN3KsUaD1ZmIA/GqcZRlySVmSpJq62LtFQ2tzHeW8c8UiyxSKHSRCGVlIyCCOCPcVNU+TKCiiigAooooAKKKKACiiigDzL9pLwm/jL4GeN9LjiaSWTTJJY1WQJuePEijJ7bkGfUcVk/sk+JH8Tfs8+CrydneaOyNozSSb2Pku0e4n3CA47AivW9QsYNUsLmzuo1ltriJoZY26MjAhgfqCa+Zf2DWbRfCPjnwZO0YuPDfiO4thGIjFII2xtZlJONxR8D0HfqfZp/vMsqw6wlGXyaaf42PNqe5i4S/mTX3ao+oqKKK8Y9IKKKKACiiigAprfw59adTJvuZzgjkUAfEX7ETq/7R3xnZCDG1xOVK9CDeyEEV9NftCeIh4Y+CfjjUfNlieLR7lUkhyHV3jKKQcjGGYHINfMn/BOHwzs1D4ieIWDANcxadEwZdpwXkcbRyCMx89Pmr1r9t64muvhPpnhu02td+JdesNLRBMI3IMm87c9T8gHPHzV9lmcY187jS6LkT+UY3PnsHJ08uc+r5mvm3YZ+wZfW9z+zboUEMqvLa3d5DMqj/VuZ2cKffa6n8a+iVr5X/Y8VfBfjz4u/DsbY7bRtb+12MPnZKQS7lAAIBICrHk9MkDivqgV4ucW/tCrKPwyfMvSWq/M9DL3/ssI9UrfNaMWiiivHPRCik3D1paACiiigAooooAKKKKACiiigDgfjp4ZPjL4R+MNFRHeW70qdY1jYKS4QsoyeBlgtcJ+xPr3/CQfs5+EJGDK9pHNZnL7t3lysM+2RjivdpIY5kZJEV0YYKsMgj0Irk/hn8LfDvwj0OXRvDFi1hp8lxJdvG0rPmR8ZILdsKAAOABXoRxEfqc8NJa8ya+5p/ocrov28ay7WZlftAapJoPwV8d38cazNFot0BGxIB3Rsv8AX9K5j9jUbP2Z/AqHqLaX/wBKJau/taapFo/7O/j2edmWOTTWtl2ru+eR1jT8NzjNaf7O67fgH8PduR/xJLQ8DOf3S1v8OVOy3qflH/gmFv8AbV5R/N/8A+Sfj58W/FvxT+Knijws0V7b/CrwxqNva+IotNdIpnt/NCSTSufm2ZDdDgBQSO9fR2t/sf8Awp8Q+HxYWnha10eWOMpaanphaK6gYkMJA4OWYHGC2cdOnFebfsz32k3nx2+P/hTVxDdXmoanJM6MgQTwB5I5VAyTgGRcj/aFfWsSwwoI0ACr8oUHpx0r080xU8JKnhsN+7UFF6O17xTu+7vdX7aHLhKEa/NWqvmcm9+lnsj5z/Z/+Kmo6D8RPEHwc8W3+oa54i0eZ5NP1eaBT9qsdiOhlZCQrAMBluuQM54r1r42TNH8GfHkkbFWXQb4gqcEEW79xXyJ8RPi1/wh/wC28/i1EkTwrpH2bw5rGox5eFfNiZsNtxgq2Tg55hJxnAr3z9rbx9D4X+BurWVvcSvqviWP+ydLitk817iWYfdAHUFNwz/tDHWjE4GX1vCzhG3tVFvspPfTppaXoxUcQvY1oyfwXXy6Fz9jXSIdH/Zr8ErCWK3FtJctvIJ3ySuzYwBxk8CvHf8AgpNqSf8ACEeDdHjE73t9qkkyRxqxVlSLZhscZ3TJgH1Poa9P/Ym8Wwa9+z7olixSK+0N5tMurfdlo2SRipYdiVIPHHbPBrwD9rT4xaT4w+LXgaTR1l1Dw/4K1qP+09bhhZ7dZ3mjcwq4G0kJCxznkgjtz24ClUefzny35JSb/G33u1u/Q5sVUh/ZsYp/Ekvyv93U+xPgj4RuPAXwl8J+Hr2aWe80/T4oZmmYsQ+0FlHAwFJ2gdgAO1dukiyZA6jqK5Tx18RNA+HPhW78R65qUVlpMCeZ52dxkJHyqg/jZjjAHrXyZof7VHi7wz8adK8SfEG2vPCvw08TadINMtbiJnS3RTmOVtoJ8xioJHPyyqQMc189QwOJzD2leK7v1e7Ue76nq1MVRwvLTk+3yXd+R9p6hq1ppNvcXN5cR2trAu+WeZwiIvqzHgD614J4g/bh8BWeqHTPDNnrXj2/BIaHw9YtKABu3YY434x/Dkc9a+Wv2hfjD4s/aW8RX/8AwiOjaxq3w30ALLPBaROq3IBJaabbyAQp2g8qBnGa+tfgT8WPgtF4HifwhqWheGbBUQz2EzpaTRNtPD7yGkIwfmy2cHBr1p5RHAYeNbEwdSUvsp25e3Pu7vorL1OKOOeJqSp0pKEV1evN3t0su5yWn/t7eFbfxZBpniXwr4k8IW8yBvterWoURZ3cvGMsFOBggHv2Ga+nLC+t9UsYbu0ljuLWdFkimjYMjqRkMCOCCOcivlX42/tIeHPiRpVx8P8A4eaVF8R/Eep/uIibTzrGzydvnu7ja23qCOBnJbjB9w+APw3l+Efwm8P+FJ7z7fcWMTGWZTlS7uXZU/2AWIHsBXn5hhqNKjCr7N0pt25W27r+bXVdtd+h0YWtUlUlDmU0uq79j0SiiivDPVCiiigBG5Uj2r5f+DE0fhv9sb4zaBG00MOpW9rq8Vu6HEr7V8yQEjpulIHOOSB0r6gb7p+lfL3j6f8A4QH9tjwPrAAitfFWh3GlXL7s5aHdICc8KABHyOuDXr5f78a9Jfag/wDyW0v0PPxfuunU7SX43X6n1Cp3AGlr5kk/ag8UfFDxFq2gfBjw3aeIW0shbrXtYuRFYqTnaUUEO4YqwHTOM9Oavv45/aH8IrJqGu+DfC/ifSoo1ea38PXkkN4BjL+WJCQ7Lz8uBnHBORUvLa0Fao4xl2ckn93+Y/rtJ6xTa7pNr7z6MoriPhn8VtC+K2gjVdEupCkbtBc2tzGYbi1lXrHLEeUb2PXqM12pb5cg15s4SpycJq0l0Z2xlGaUou6Y6ioftCZGHDZ9DU1Z3uUFcV8XviJYfDH4c674lvpAsdhbs8SsP9ZMfliQDvucqPxPpXZs3OM818b/ABJvZ/2qfiFq+kafM0vwx8EQTXOoXcTny9T1ARtsjRhwyoQB7bWIPzKa9PAYaOIq3qu1OOsn5dvV7Jf8E4sVWdOHLD4paL+vI7j9grRZNP8A2e7K/kMTPqmpXd7+7HQeZ5eDkdcxnHsRUnx48rxb+0V8E/CRELrbXV14guI5Q24LDH+7wRx95W49VGffT/YeJ/4Zc8HHjOLvt/09S9qoaHI3iT9tjxFN9oma38N+FreyEez92slxKJCM+u3aR/wL0r16zazTGVX9nn/NxX5o4o2eEoU7fFy//Jfoc38SJH+Dv7Yvg/xmxcaF41tv7CvfLVmAuFCrGTg45Ihx7K/HevqxTnNeMftSfCW6+LXwpurTSm8nX9LmXVNLl/iE0QJ2A9QWG4A54O3PSvOLn9saKH9mGLxZblH8bORo7ac/ysuo7cM+0jO0D96BjB+7nrXPKhUzKhQnR1nH92//AG1vytdeXKaRqRwlWoqmkX7y/X8T0/40ftGaL8LZ7PR7OxuPFfjG8fba+HNLIa4PGd8mM7F5HOCSMkAgHHH+Hde/aU8WSNcXOkeC/BumyYKQaiJrqdBkAgiOTBPU849Ditr9mX4Ap8NdIk8SeJQ2qfELXF8/VdSumMkiFyG8pWIGAOM8ckY6AY9wkWJc5IB9zXNVrYfCXo0YKbW8nrd/3Ve1vVa+RrTp1a6VWpJx7RX66bnzPr37Rnjz4S+JLdPiX4HtYfCc8q2Y8TaBcGaDzG2hXZH5RMluGweOM9/pyFg0MZB3AqDn14r5W/ap1aT4v67onwS8OBbrVb+5jvdZu0BZNMtYzuBcj7rMcYBzxgcbgR9S2NtHZ2NvbxLtihjWNFxjCgADj8KWMp01Ro1FBQlJO6V9vsy1el9fIrDyk6lSPNzRT38+xPRRRXkHeFFFFABRRRQAUUUUAFFFFAHh37acL3X7NfjaNCoKw28nOei3MTHt6D/HHWuq/Z0Xd8A/h8P+oHaf+ilqn+1Aw/4Z9+IQ7/2NP/6DT/2X9Tg1b9nvwBPbFmjXSYYCWGPnjHlv/wCPIa9q7eVKPRVH+MV/keda2Nb68q/M8K+Jn7NXjzSfjh4i+Jnw6vre21bfDf2lpKqLDd7lKXUDkkYZtqsMja28/MGANauqfHv45axa/wBi6T8F7rS/EDb45NSurrfZxY4Do3yqxwe7n8elfVxQZJ2jnrSeWvXaufpR/ajnGKxFKNRwVot30tsnZq69fvF9RjFy9lNx5t7W3+f6Hg3w5/Zc0fRPgzqfgvxXJ/wkd1r87X+s3uShkum2neh6goQMN1JGT1wOZ+HP7IOq6D4w8N3vizx7c+LtB8LsZND0mS3MYt23ZjZyWO7bwRx1VcYCgV9QBRjlRRg+lYrNMWnO0vj320uraaaaaaW0NXgqDUVb4f6+Z8y+OP2L11bxlreq+FfHeseCdN8QOH1bS9LULFKQDyu1l4LEnDAgbmx1rv8AXf2ZfBuqfBNvhrbWh03RQitDNCoM0dwDkXBJHzOT1z1BI4B49cx6ig9OmamWZYuagnN+5Zr5betul7jWEoRcmo77/Pf0+R8n+E/2JJrjxBpNz8RfG95480fRFWPS9JmVkhjVTgK+5mJXAHAxnGCSBivo/wAWeCdH8YeH59G1bT7e906aJofJkiUhAUK5XI+UgHgjBFdAV9gaXGRyM1niMfiMVOM6stY7WskvRKy16jpYWlRi4Rjvv1/M+ff2efh/4w+CWmx+BJ9Cs9R8OQ39xJF4kt79Y3kicblZ7YqSXz8pw2O9afiT9jH4S+KteXVrrwwttc+b5skdjPJBDMxxnfGp24OOduM5PrXtwQdcCnU/7QxPtZV4zcZS3abV/NjWFoqCpyimltdbHL+Cfhl4Z+G+kjTPDGj2mi2WQWjto8GRh/E7HLOfdiTXSxx+X0JP1p9FcU5SqSc5u7fU6IxUVyxVkFFFFSUFFFM85d2MN+VADz0r4x/4KR6XcweFfBWu2imGS2v57RryM7XQSxZCZByQ2xuOnHNe4ePv2rvhh8O7+TTtS8TR3OqqQhsdMie7mDE424QEBs/wk56cc18y/tPfHqX4v/DO7trL4YeJbbSNNvYL1ta1mBoIYGU7WV1XoGVyud38fYgV9XkWFxNLHUcQ6b5L9dE09Hu1fc8TMq1KeHqUlL3reuvyPo/9lHwz4b0P4I+Errw9bWsR1CwhmvLqGBI5LicLhzJt6sDuXk9q9ieENkEn8zX5kaB8ftX+AOrTal8PNI1i08E6oVuJdD8S2LC2t5mUEiCYPllOWw3BIC5Dda918FftifE34nWMk3hT4X6Zq7AN8sXiGLzEx3aI4bHTsOo55zV5hkWMlVlitHCT3bSav0ld/wDAMcLmWHUFTas12Td/Sx3P7QljP8G9esfi/wCHZjYxR3dva+KbCPCR6nauwjWVwfl82Ivw2M4PUAV0XxW/aj8O/DprfR9Lifxj4wvkH2PQtIHmyOzD5GkK5CKeP9rBzjHNeJ/GTwh8avGnwz8S698R/E2m+DvDen6fJNJ4e0ENMbvahwJW3H5S5UEFmxyQBgZ9n/Zb+FHhbwZ8LfCutaToltDrWraRbXV5qLZeeaSSJWf52yQuT90YAA6VnVo4ejhqdbEP2kotxtF6PS6Tl5eV9GldaGtOpWnXlTpx5E0nrvvrZefmec/sz+OviL4l/aD8d6f8QNRlsr2x06N4fDMLZtrdXkVspt+UlFKqXJJbzOvFfTfirxto3gXQ59X8Qaja6TpluuZLi6k2qMnAA9STwAOSeleQfF79l6fxz4sl8Z+FPF+q+CPGMlotnNeWUpMVzGv3Q6gggDjof4RwTzWb4T/Y3sZL+DVfiT4l1T4napEo8uLVpm+xwttAJWIk7uR/EcEYyOuefESwOLccS5qGiThGLvdLW3Sz7t3NKSxNBOko8z195vT59fuPNfEX7YGq/H5rr4f/AA48Oahpus6tcfYo9amkGILNsiS44wY224I5PBPO7aD9E6D8MdI+DfwPv/DWjg/ZLTS7kyTOAHuJTExeV/8AaY/lwO1eg6fo1lppZ7aygtZGRUZoo1UlVGFUkDkKOAO3auX+NU1jZ/CHxpNqULXFgmj3bTxIcM6+S2QOR/OuariqdaUKGFp8kLptXu2+7em3Reb7nRGjKEZVK0ryt6WPOv2HVP8Awy/4OHQ4u+Dxj/SpuDVf9nRTrvxT+N/iQXcl1DPr8WmQuE2pstYiuAcDJUuVP+771ofshpD4Z/Zc8IzT7lgjsp7x8ZY7WlkkOMqvY9MfQnqYv2L7aWT4Ixa1OZnuvEGqX2ryvcKFL+ZOwVgABwyorf8AAjjjArsxkuWeOqfzT5fvk5f+2nNh1zRw8O0b/gl+p7t5e5cZxX55/tHeB5fC/wC2To9z4K8NQ6xqclvD4im0hpAFvJ0eVpSFZh8xWIMAv8QJwa/Q+vnf9qXwnr+ma94N+KfhSxbVNV8ITSNe6fEpaW5sZBiQIOhYLv4HPzE87cVhkeK+q4lpfbjKOr0ba0vquvU0zKl7ain/ACtMpL+3T4S0f9x4t8MeKvCOpr8ps73TS+WHD7WB+ZVPGcD6Vmax8d/iJ8dLdtP+EPhm80LT5Ttl8XeIofIiVSMhreMhi5wQc4bg/dBwa988BeO9A+JXhy113Q76DVNPnUANGwcxvgEow/hYZ5U106srBduNo6Y6Vk8Rh6Evdw9prvK6X/btvzZoqVaotat4+SS/G55X8DvgBovwa0++lgnn1TxJqm2XV9bumLTXcvJPU/Ku4k7eeuSSa9YX7opFHtinV5dWrOtN1Kju2d1OnGlFRitAooorIsKKKKACiiigAooooAKKKKAOT+KfhUeNfh34n0P7OLqTUNNuLaOFn2B5GjOwFuw3ba5j9mHwfq/gH4G+FPD+vW32TVrG3dLiDerlGaaRgCVOOjDp616nRXR9Yn7B4f7Lal80mv1MfZR9r7XrawUUUVzmwUUUUAFFFFABRRRQAUUUUAFFFFABRRTX5Uj1oA5T4nfErQ/hL4RvfEviC6Ftp1ovIUZklc/djQd2Y8D8c8CvnvS9D+Jn7VMMGr69qN18OPhtdAtBomnuU1K/hx8rSyY+VTweMAj+Hoxl+IFmfjZ+11o3gnUgr+F/Bmnrrd1YyYZb6d9gTcuMFRuQYPo396vqiPA28Y4xxXt80ctpQlCP72a5rvXlT2strve/TSx5lnjJyu/ci7WXV9b+RwHw5+A/gT4V2+PDnhy1srjOWvJFMtw5xjJkclvwBxzU/wAcND/4SX4O+M9LRYnkudIuUQXHKBvLYgng9CAenau7PQ1VvreK8tZLaaNJoZlMckcnKsrDBB9QQcY968v29SVVVqknKSad3rsdnsYKm6cVZW6Hjn7KepQ+Ov2afBhv4Fu4TYNYyxXmJRIIneLkHPBC9D2OKq+Pv2P/AIeeLr59U0+xuPCOvEZTVPD0zWzo2MZ2r8v5AZ9awv2FLqS1+GGveGZzIbjw34ivbBhj92q7gwEZ/u5Zj+Jr6UX7tepjK1bCY6s6MnG8m9NNH+nkceHp06+Gp+0inoj5W8Q/sv8AxU17QX8LXnxsutR8JXEQjvI7zS0a6kXdkgPksRwMZftivpTwzodr4Z0HTtJsoxFZ2NtHawRqMBURQq/oBWtRXDiMZWxMVGo1Za6JLV7vRLXTc6qWHp0W5R3fd3/MKKKK4jpCuN+M2jP4i+EnjLTI5RBJd6PdwrIV3BSYWGcd67KsPxyC3gvX1UFmbT7gKoBJJ8puK2oycakZLo0Z1FzQkvI+a/hL4yn0n/gn/Jq810sE1noV/BBMyA+WVkmihTGOSDsXn8a9j/Zs0d9B+BHgOxkkErJo9u+8LtBDoHHHtux+FfKGla7HJ/wTpsNJtb22iub7UP7Gl87PytLfFmAHHzBCG47Zr7r0eyTTdLtLKMKsdrCkACKAvyALwB06dK+gzePJ7WK+1Vn90dP/AG5nl4H3pQfaEfxv/kaNRPErSZy2cY4NS0V8yeweK+JP2ZdIj1658QeBtZ1L4ceILtt9zcaGVNtdNkkGa2cGN/mJPbPeun8F6P8AETRdYEOu+JdG8R6Iq8TLpj2t8W2gDO2Qx4yCT8uTntXodFdMsTVqR5Jvm9Um16O1/wATCNGEXeOnzdvuEXPeloormNwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigApkmdpx17U+k25pMD5gmWXwF+3R9pv9v2Lxr4fFrZSZPyzW5RmQ4HUiLjkffHfNfTkQ69ce9eP/tOfDC9+IHgm1vvD7CDxl4cuV1bRbgD5vNQ5aLIHR1GMdCwWt34F/GK0+NHgi31qKA6fqMTta6npr532d0nDxnIB7gj2PrmvYxV8Rh6eJS+FKEvK3wv0a09UefRtRrSovrqvnv8Aiej00r0I60ueKazEdOvvXjXO8+cf2O7CXTNQ+MVtcxeTcL40vGZGIzyFIP5EH8a+kelfOP7LXlal8SPjjrdtKJbW68VNbpmIqwMSYY5POCWwBj+EnJzX0aK9bM/97l6R/wDSUcWC/gJeb/Ni0UUV5Z3BRRRQAVHcKXUAZ6+tSUyR/LAJ6ZovbUW6sfn/APDDTRdfC/wz4ReFbu/0n4vR293bsAUkC72kxn7yhAxPsK/QLy1XoK+D/hfLY+Ff2ntR8Ga4kb3cHju41iwzGwnf7TZz7XUhtojA8vOeTvHHBx9419Lnjft4ro/eXnzankZalyN9Vp9wUUUV80ewFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAMkTcc47V83fGT4b678IfE+ofGD4brvuCPO8TeGsnydWhXlpUX+GZRk5A55PXIb6UpsiCRcGurD4iWFm5R1T0a6Ndn/W+phVpKrGz0fR9mcr8OfiNonxN8H6d4i0O7S8sL2MP8jhjE2BujfHRlPBFbmqapbadp9xe3Mhit7VGmlfH3VUEkn8Aa8i8a/s7JJfahq/gHxHqXw+12+cS3DaexeyuW27WaW1zs3bc/Mu1snOTXmy/Bv8AaB8VabL4I8XeONHvPBl6ixahq9vEf7Qkgz+8hT5F+8OCW7ADJGQ3ZHCYas3UjWUY32lfmS67Jp2+V+xzOtWglGVNttbrY6n9iW0N18Ode8Utby23/CVeIr7VI45XDYjL7VwMcfdYc+me9fRorI8K+G9O8I+H9M0bSrf7Jp2n26W1vCP4UVQAPc8fnk1rDvXJi66xNedVKyb09On4HTQpulSjB7oWiiiuQ3CiiigApGGR60tFAHhPiT9mG31n9pLQfipa6mtmLNA95YeVuNxMsZjR1bOFGwgHj+AY6mvdqKK6a2Jq4hQVWV+RWXkl0MadKFJycFa7uwooormNgooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD//Z"}, {"name": "Arch. Riccardo Hoops", "discipline": "CAM e DNSH", "role": "Ispettore Tecnico", "signature": "data:image/png;base64,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"}, {"name": "Arch. Bianca Barbieri", "discipline": "CAM e DNSH", "role": "Ispettore Tecnico", "signature": "data:image/png;base64,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"}, {"name": "Ing. Marcello Caccialupi", "discipline": "Acustica", "role": "Ispettore Tecnico", "signature": "data:image/jpeg;base64,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"}, {"name": "Geom. Massimo Tamberi", "discipline": "Sicurezza, Cantierizzazione e relativa documentazione economica", "role": "Ispettore Tecnico", "signature": "data:image/png;base64,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"}] as const;


type SignatureDatabaseEntry = (typeof SIGNATURE_DATABASE)[number];

function signatureKey(entry: SignatureDatabaseEntry) {
  return `${entry.role}||${entry.name}||${entry.discipline}`;
}

function getSignatureEntryByKey(key = "") {
  return SIGNATURE_DATABASE.find((entry) => signatureKey(entry) === key);
}

function getResponsabileOptions() {
  return SIGNATURE_DATABASE.filter((entry) => entry.role === "Responsabile Tecnico" || entry.role === "Coordinamento");
}

function getIspettoreOptions() {
  const allowedInspectors = new Set([
    "Arch. Veronica Laino",
    "Arch. Arianna Brunetti",
    "Ing. Salvatore Grimaldi",
    "Ing. Bruno Gabrielli",
    "Ing. Carlo Renda",
    "Ing. Gianluca Biaggioli",
    "Ing. Marta Dominijanni",
    "P.I. Mauro Garofalo",
    "Arch. Riccardo Hoops",
    "Ing. Marcello Caccialupi",
    "Geom. Massimo Tamberi",
  ]);

  return SIGNATURE_DATABASE.filter(
    (entry) => entry.role === "Ispettore Tecnico" && allowedInspectors.has(entry.name.trim())
  );
}


function addPdfPageNumber(doc: jsPDF) {
  const pageCount = doc.getNumberOfPages();
  const pageWidthCurrent = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  doc.setFontSize(8);
  doc.setTextColor(100);
  doc.text(`Pagina ${pageCount}`, pageWidthCurrent - 10, pageHeight - 6, { align: "right" });

  // Codice qualità aziendale ITS - fisso per tutti i progetti.
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  doc.setTextColor(110, 110, 110);
  doc.text("Mod.PR 09.1.TR.rev.1 19.06.2026", 6, pageHeight - 38, { angle: 90 });
}

function formatCommentDateIt(value = "") {
  const raw = String(value || "").trim();
  if (!raw) return "";

  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return `${match[3]}-${match[2]}-${match[1]}`;

  const fallback = new Date(raw);
  if (!Number.isNaN(fallback.getTime())) {
    return fallback.toLocaleDateString("it-IT").replace(/\//g, "-");
  }

  return raw;
}

function normalizeCommentAuthor(author = "") {
  return cleanPdfText(author).replace(/\s+/g, " ").trim();
}

function getInitialsFromName(value = "") {
  const cleaned = cleanPdfText(value)
    .replace(/\([^)]*\)/g, " ")
    .replace(/\b(arch|architetto|ing|ingegnere|geom|geometra|dott|dottore|dottssa|dott\.ssa|avv|prof)\.?\b/gi, " ")
    .replace(/[^A-Za-zÀ-ÖØ-öø-ÿ\s'-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!cleaned) return "";

  const parts = cleaned.split(" ").filter(Boolean);

  if (parts.length === 1) {
    return parts[0].slice(0, 2).toUpperCase();
  }

  return `${parts[0][0] || ""}${parts[parts.length - 1][0] || ""}`.toUpperCase();
}

function getRedattoreFromRow(row: any) {
  const createdBy =
    row?.creatoDa ||
    row?.["Created by"] ||
    row?.createdBy ||
    row?.ispettore ||
    row?.Owner ||
    "";

  return getInitialsFromName(createdBy);
}

function isSolibriCheckingRow(row: any) {
  const source = String(row?.sourceFile || row?.fileName || "").toLowerCase();
  const origine = String(row?.origine || "").toLowerCase();
  const tipoVerifica = String(row?.tipoVerifica || "").toLowerCase();

  return (
    Boolean(row?.isSolibriChecking) ||
    source.includes("solibri") ||
    origine.includes("solibri") ||
    tipoVerifica.includes("checking")
  );
}

function getDisciplinaDisplay(row: any) {
  const rawDisciplina = String(row?.disciplina || "").trim();
  const isAccountDiscipline = /@/.test(rawDisciplina);

  return isSolibriCheckingRow(row) || isAccountDiscipline ? "BIM" : row?.disciplina || "Non assegnata";
}

function isBimModelRow(row: any) {
  const elaborato = String(getElaboratoKey(row) || "").toLowerCase();
  const rawDisciplina = String(row?.disciplina || "").trim();

  return (
    isSolibriCheckingRow(row) ||
    /@/.test(rawDisciplina) ||
    elaborato.includes(".ifc") ||
    Boolean(row?.isBimModel)
  );
}

function getOggettoVerificaTipo(row: any) {
  return isBimModelRow(row) ? "modello" : "elaborato";
}

function getRedattoreFromComments(comments: any[] = []) {
  const ispAuthors = comments
    .filter((c: any) => String(c?.role || "").toUpperCase() === "ISP")
    .map((c: any) => normalizeCommentAuthor(c?.author || ""))
    .filter(Boolean);

  return Array.from(new Set(ispAuthors)).join(" / ");
}

function commentsToPdfRichText(comments: any[] = []) {
  if (!Array.isArray(comments) || comments.length === 0) return "";

  return comments
    .map((c: any) => {
      const role = c.role ? `[${String(c.role).toUpperCase()}] ` : "";
      const author = normalizeCommentAuthor(c.author || "Autore non indicato");
      const date = formatCommentDateIt(c.date || "");
      const header = [role + author, date].filter(Boolean).join(" - ");
      const comment = cleanPdfText(c.comment || "");
      return [header, comment].filter(Boolean).join("\n");
    })
    .join("\n\n");
}

function exportDetailPdf(rows: any[], title = "", headerData: PdfHeaderData = {}) {
  const templateMode = Boolean((headerData as any)?.templateMode);
  const sourceRowsForSummary = Array.isArray(rows) ? rows : [];

  if (!templateMode && (!rows || rows.length === 0)) return;

  if (templateMode) {
    rows = [];
  } else {
    const printableRows = prepareRowsForPdfExport(rows).filter((r: any) => {
      const tipo = cleanPdfText(r.tipo || r["Tipologia rilievo"]).toUpperCase();
      return tipo === "NC" || tipo === "OSS" || tipo === "DA NC A OSS";
    });

    if (printableRows.length === 0) return;

    rows = printableRows;
  }

  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const generatedAt = new Date().toLocaleString("it-IT");
  const disciplina = detectDisciplinaForPdf(rows, title);
  const pageWidth = doc.internal.pageSize.getWidth();
  const marginX = 10;
  const headerY = 8;
  const headerW = 273;
  const responsabileEntry = getSignatureEntryByKey(headerData.responsabileTecnicoKey || "");
  const inspectorEntries = (headerData.ispettoreKeys || [])
    .map((key) => getSignatureEntryByKey(key))
    .filter(Boolean) as SignatureDatabaseEntry[];
  const inspectorRows = Math.max(1, inspectorEntries.length);
  const signatoryH = inspectorRows > 12 ? 7 : inspectorRows > 8 ? 8 : 10;
  const headerBoxH = 68 + 10 + signatoryH * (1 + inspectorRows);
  // Testata solo sulla prima pagina, con colori e logo ricavati dal template Excel ITS.
  doc.setFillColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.rect(marginX, headerY, headerW, 3, "F");

  doc.setDrawColor(210, 210, 210);
  doc.setLineWidth(0.25);
  doc.rect(marginX, headerY + 3, headerW, headerBoxH);

  doc.setFillColor(255, 255, 255);
  doc.rect(marginX, headerY + 3, 66, 17, "F");
  try {
    doc.addImage(ITS_LOGO_BASE64, "PNG", marginX + 4, headerY + 5, 35, 14);
  } catch (error) {
    doc.setTextColor(30, 117, 192);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("ITS", marginX + 5, headerY + 13);
  }

  doc.setFillColor(255, 255, 255);
  doc.rect(marginX + 66, headerY + 3, 142, 17, "F");
  doc.setTextColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(17);
  doc.text("SCHEDA DI ISPEZIONE", marginX + 137, headerY + 12, { align: "center" });

  doc.setFontSize(10);
  doc.text((disciplina || "").slice(0, 90), marginX + 137, headerY + 17, { align: "center" });

  doc.setFillColor(ITS_LIGHT_BLUE[0], ITS_LIGHT_BLUE[1], ITS_LIGHT_BLUE[2]);
  doc.rect(marginX + 208, headerY + 3, headerW - 208, 17, "F");
  doc.setTextColor(80, 80, 80);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  doc.text(headerValue(templateMode ? "" : headerData.codiceScheda || "").slice(0, 55), pageWidth - marginX - 4, headerY + 12, { align: "right" });

  // Sezione 1 - Dati commessa
  doc.setFillColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.rect(marginX, headerY + 20, headerW, 7, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.text("DATI COMMESSA", marginX + 1, headerY + 25);

  const leftX = marginX;
  const labelW = 68;
  const midX = marginX + 168;
  const row1Y = headerY + 27;
  const rowH = 8;

  function labelCell(label: string, x: number, y: number, w: number) {
    doc.setFillColor(ITS_GRAY[0], ITS_GRAY[1], ITS_GRAY[2]);
    doc.rect(x, y, w, rowH, "FD");
    doc.setTextColor(80, 80, 80);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7.5);
    doc.text(label, x + w - 2, y + 4.7, { align: "right" });
  }

  function labelTallCell(label: string, x: number, y: number, w: number, h: number) {
    doc.setFillColor(ITS_GRAY[0], ITS_GRAY[1], ITS_GRAY[2]);
    doc.rect(x, y, w, h, "FD");
    doc.setTextColor(80, 80, 80);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7.5);
    doc.text(label, x + w - 2, y + Math.min(h - 2, 4.7), { align: "right" });
  }

  function valueCell(value: string, x: number, y: number, w: number) {
    doc.setFillColor(255, 255, 255);
    doc.rect(x, y, w, rowH, "D");
    doc.setTextColor(35, 35, 35);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.text(headerValue(value).slice(0, Math.max(15, Math.floor(w / 2.1))), x + 2, y + 4.7);
  }

  function valueMultiCell(value: string, x: number, y: number, w: number, h: number, maxChars = 300) {
    doc.setFillColor(255, 255, 255);
    doc.rect(x, y, w, h, "D");
    doc.setTextColor(35, 35, 35);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.2);

    const safeValue = headerValue(value).slice(0, maxChars);
    const lines = doc.splitTextToSize(safeValue, w - 4).slice(0, Math.max(1, Math.floor((h - 2) / 3.4)));
    doc.text(lines, x + 2, y + 4.2);
  }

  function signatureCell(value: string, imageDataUrl: string | undefined, x: number, y: number, w: number, h = rowH) {
    doc.setFillColor(255, 255, 255);
    doc.rect(x, y, w, h, "D");

    if (imageDataUrl) {
      try {
        const props = doc.getImageProperties(imageDataUrl);
        const ratio = props.width && props.height ? props.width / props.height : 3.2;
        const maxW = Math.min(w - 4, pageWidth - x - 12);
        const maxH = h - 2;
        let imageW = maxW;
        let imageH = imageW / ratio;

        if (imageH > maxH) {
          imageH = maxH;
          imageW = imageH * ratio;
        }

        doc.addImage(imageDataUrl, getImageFormatFromDataUrl(imageDataUrl), x + 2, y + Math.max(1, (h - imageH) / 2), imageW, imageH);
        return;
      } catch (error) {
        // In caso di immagine non valida, usa il testo di fallback.
      }
    }

    doc.setTextColor(35, 35, 35);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.text(headerValue(value).slice(0, Math.max(15, Math.floor(w / 2.1))), x + 2, y + Math.min(h - 2, 4.7));
  }

  function drawSmallSignature(
    imageDataUrl: string | undefined,
    x: number,
    y: number,
    w: number,
    h: number,
    signatoryName = ""
  ) {
    if (!imageDataUrl) return;

    try {
      const props = doc.getImageProperties(imageDataUrl);
      const ratio = props.width && props.height ? props.width / props.height : 3.2;
      const isBiaggioli = signatoryName.toLowerCase().includes("gianluca biaggioli");
      const maxW = w - (isBiaggioli ? 1.5 : 4);
      const maxH = h - (isBiaggioli ? 1.2 : 2.5);
      let imageW = maxW;
      let imageH = imageW / ratio;

      if (imageH > maxH) {
        imageH = maxH;
        imageW = imageH * ratio;
      }

      doc.addImage(
        imageDataUrl,
        getImageFormatFromDataUrl(imageDataUrl),
        x + (w - imageW) / 2,
        y + Math.max(0.8, (h - imageH) / 2),
        imageW,
        imageH
      );
    } catch (error) {
      // Salta immagini non valide.
    }
  }

  function signatoryRow(label: string, name: string, signature: string | undefined, x: number, y: number, w: number, h: number) {
    const signatureW = 82;
    const nameW = w - labelW - signatureW;

    labelTallCell(label, x, y, labelW, h);

    // Nome sempre a sinistra.
    doc.setFillColor(255, 255, 255);
    doc.rect(x + labelW, y, nameW, h, "D");
    doc.setTextColor(35, 35, 35);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(h < 8 ? 6.6 : 7.5);
    doc.text(
      headerValue(name).slice(0, Math.max(15, Math.floor(nameW / 2.1))),
      x + labelW + 2,
      y + Math.min(h - 2, 5.2)
    );

    // Firma sempre a destra, dopo il nome.
    doc.setFillColor(255, 255, 255);
    doc.rect(x + labelW + nameW, y, signatureW, h, "D");
    drawSmallSignature(signature, x + labelW + nameW, y, signatureW, h, name);
  }

  const valueW = headerW - labelW;
  const oggettoRowH = 18;
  const notaY = row1Y + rowH + oggettoRowH;
  const dataEmissioneY = notaY + rowH;
  const firstSignatoryY = dataEmissioneY + rowH;

  labelCell("Committente / Stazione Appaltante:", leftX, row1Y, labelW);
  valueCell(templateMode ? "" : headerData.committente || "", leftX + labelW, row1Y, valueW);

  labelTallCell("Oggetto del Contratto / Accordo Quadro:", leftX, row1Y + rowH, labelW, oggettoRowH);
  valueMultiCell(templateMode ? "" : headerData.oggetto || "", leftX + labelW, row1Y + rowH, valueW, oggettoRowH, 300);

  labelCell("Nota di Ricezione Elaborati e data:", leftX, notaY, labelW);
  valueCell(templateMode ? "" : headerData.notaRicezione || "", leftX + labelW, notaY, valueW);

  labelCell("Data emissione:", leftX, dataEmissioneY, labelW);
  valueCell(templateMode ? "" : headerData.dataEmissione || "", leftX + labelW, dataEmissioneY, valueW);
  signatoryRow(
    "Responsabile tecnico:",
    templateMode ? "" : headerData.responsabileTecnico || "",
    templateMode ? "" : responsabileEntry?.signature || headerData.firmaResponsabileImage || headerData.firmaImage,
    leftX,
    firstSignatoryY,
    headerW,
    signatoryH
  );

  const inspectorsToPrint = templateMode
    ? [{ name: "", signature: "" } as any]
    : inspectorEntries.length > 0
      ? inspectorEntries
      : [{ name: headerData.ispettore || "", signature: (headerData.firmaIspettoreImages || [])[0] || "" } as any];

  inspectorsToPrint.forEach((entry: any, index: number) => {
    signatoryRow(
      index === 0 ? "Ispettore:" : "",
      entry.name || "",
      entry.signature || "",
      leftX,
      firstSignatoryY + signatoryH * (index + 1),
      headerW,
      signatoryH
    );
  });

  // Sezione 4 - Rilievi, coerente con il template.
  const sectionY = headerY + headerBoxH + 30;
  doc.setFillColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.rect(marginX, sectionY, headerW, 7, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.text("RILIEVI ITS CONTROLLI TECNICI   NC = Non conformità   |   OSS = Osservazione", marginX + 1, sectionY + 5);

  doc.setTextColor(80, 80, 80);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7.5);
  if (!templateMode) doc.text(`Numero righe: ${rows.length}`, pageWidth - marginX, sectionY - 2, { align: "right" });

  const tableRows = rows.map((r: any, i: number) => {
    const allComments = Array.isArray(r.comments) ? r.comments : [];
    return [
      String(i + 1),
      cleanPdfText(r.id),
      cleanPdfText(r.tipo),
      cleanPdfText(getDisciplinaDisplay(r)),
      cleanPdfText(getRedattoreFromRow(r)),
      cleanPdfText(getElaboratoKey(r)),
      cleanPdfText(r.descrizione),
      cleanPdfText(commentsToPdfRichText(allComments)),
      cleanPdfText(translateStatus(r.stato)),
    ];
  });

  autoTable(doc, {
    startY: sectionY + 7,
    head: [["N.", "ID rilievo", "Tipologia", "Disciplina", "Redattore", "Elaborato", "Descrizione", "Gestione rilievo", "Stato"]],
    body: tableRows,
    theme: "grid",
    tableWidth: headerW,
    styles: {
      font: "helvetica",
      fontSize: 7,
      cellPadding: 1.8,
      overflow: "linebreak",
      valign: "top",
      lineColor: [200, 200, 200],
      lineWidth: 0.1,
    },
    headStyles: {
      fillColor: ITS_DARK_BLUE,
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
    alternateRowStyles: { fillColor: ITS_TABLE_LIGHT },
    columnStyles: {
      0: { cellWidth: 9, halign: "center" },
      1: { cellWidth: 21 },
      2: { cellWidth: 21 },
      3: { cellWidth: 23 },
      4: { cellWidth: 26 },
      5: { cellWidth: 34 },
      6: { cellWidth: 52 },
      7: { cellWidth: 68 },
      8: { cellWidth: 19 },
    },
    margin: { left: 10, right: 10, top: 10, bottom: 12 },
    didParseCell: (data: any) => {
      if (data.section === "body" && data.column.index === 7) {
        data.cell.styles.fontStyle = "normal";
        data.cell.styles.fontSize = 7;
        data.cell.styles.cellPadding = 2;
        data.cell.styles.overflow = "linebreak";
      }
    },
    didDrawPage: () => addPdfPageNumber(doc),
  });


  const summaryRows = buildElaboratiDisciplinaSummary(sourceRowsForSummary);
  const lastAutoTable = (doc as any).lastAutoTable;
  let summaryStartY = (lastAutoTable?.finalY || 20) + 10;

  if (summaryStartY > doc.internal.pageSize.getHeight() - 50) {
    doc.addPage();
    summaryStartY = 18;
  }

  doc.setTextColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("ELENCO ELABORATI DELLA DISCIPLINA E CONTEGGIO NC/OSS", 10, summaryStartY);

  autoTable(doc, {
    startY: summaryStartY + 4,
    head: [["N.", "Elaborato", "NC", "OSS", "Esito"]],
    body: summaryRows,
    theme: "grid",
    styles: {
      font: "helvetica",
      fontSize: 7.5,
      cellPadding: 2,
      overflow: "linebreak",
      valign: "top",
      lineColor: [200, 200, 200],
      lineWidth: 0.1,
    },
    headStyles: {
      fillColor: ITS_DARK_BLUE,
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
    alternateRowStyles: { fillColor: ITS_TABLE_LIGHT },
    columnStyles: {
      0: { cellWidth: 10, halign: "center" },
      1: { cellWidth: 95 },
      2: { cellWidth: 16, halign: "center" },
      3: { cellWidth: 16, halign: "center" },
      4: { cellWidth: 55 },
    },
    margin: { left: 10, right: 10, top: 10, bottom: 12 },
    didDrawPage: () => addPdfPageNumber(doc),
  });

  const rowsWithImages = rows.filter((r: any) => Boolean(r.snapshotDataUrl));

  if (rowsWithImages.length > 0) {
    const lastSummaryTable = (doc as any).lastAutoTable;
    let imagesStartY = (lastSummaryTable?.finalY || summaryStartY) + 12;
    const pageHeight = doc.internal.pageSize.getHeight();
    const imageMaxW = 135;
    const imageMaxH = 78;
    const imageX = 10;
    const textX = imageX + imageMaxW + 8;
    const contentW = pageWidth - textX - 10;

    if (imagesStartY > pageHeight - 45) {
      doc.addPage();
      imagesStartY = 18;
    }

    doc.setTextColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("IMMAGINI NC/OSS", 10, imagesStartY);

    let currentY = imagesStartY + 8;

    rowsWithImages.forEach((r: any, index: number) => {
      const requiredHeight = imageMaxH + 22;

      if (currentY > pageHeight - requiredHeight - 12) {
        doc.addPage();
        currentY = 18;
      }

      const idRilievo = cleanPdfText(r.id || r.idTodo || `Rilievo ${index + 1}`);
      const tipologia = cleanPdfText(r.tipo || "");
      const disciplina = cleanPdfText(r.disciplina || "");
      const elaborato = cleanPdfText(getElaboratoKey(r));
      const descrizione = cleanPdfText(r.descrizione || "");

      doc.setFillColor(ITS_DARK_BLUE[0], ITS_DARK_BLUE[1], ITS_DARK_BLUE[2]);
      doc.rect(10, currentY, pageWidth - 20, 7, "F");
      doc.setTextColor(255, 255, 255);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(8);
      doc.text(`Immagine ${index + 1} - ${tipologia || "Rilievo"} - ${idRilievo}`, 12, currentY + 4.8);

      currentY += 10;

      try {
        const props = doc.getImageProperties(r.snapshotDataUrl);
        const ratio = props.width && props.height ? props.width / props.height : 1.6;
        let imageW = imageMaxW;
        let imageH = imageW / ratio;

        if (imageH > imageMaxH) {
          imageH = imageMaxH;
          imageW = imageH * ratio;
        }

        doc.addImage(
          r.snapshotDataUrl,
          getImageFormatFromDataUrl(r.snapshotDataUrl),
          imageX,
          currentY,
          imageW,
          imageH
        );

        doc.setDrawColor(200, 200, 200);
        doc.rect(imageX, currentY, imageW, imageH);
      } catch (error) {
        doc.setDrawColor(200, 200, 200);
        doc.rect(imageX, currentY, imageMaxW, 22);
        doc.setTextColor(120, 120, 120);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(8);
        doc.text("Immagine BCF non visualizzabile nel PDF", imageX + 3, currentY + 12);
      }

      doc.setTextColor(35, 35, 35);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(8);
      doc.text("ID rilievo:", textX, currentY + 4);
      doc.text("Tipologia:", textX, currentY + 12);
      doc.text("Disciplina:", textX, currentY + 20);
      doc.text("Elaborato:", textX, currentY + 28);
      doc.text("Descrizione:", textX, currentY + 40);

      doc.setFont("helvetica", "normal");
      doc.text(idRilievo.slice(0, 75), textX + 24, currentY + 4);
      doc.text(tipologia.slice(0, 40), textX + 24, currentY + 12);
      doc.text(disciplina.slice(0, 40), textX + 24, currentY + 20);

      const elaboratoLines = doc.splitTextToSize(elaborato || "Elaborato non identificato", contentW);
      doc.text(elaboratoLines.slice(0, 2), textX + 24, currentY + 28);

      const descrizioneLines = doc.splitTextToSize(descrizione || "", contentW);
      doc.text(descrizioneLines.slice(0, 5), textX, currentY + 46);

      currentY += imageMaxH + 22;
    });

    addPdfPageNumber(doc);
  }

  doc.save("Dettaglio_selezione.pdf");
}
function toDashboardExportRows(rows: any[]) {
  return rows.map((r: any) => ({
    ID_Rilievo: r.id || "",
    "Tipologia rilievo": r.tipo || "",
    Disciplina: getDisciplinaDisplay(r),
    Redattore: getRedattoreFromRow(r),
    Elaborato: getElaboratoKey(r),
    Descrizione: r.descrizione || "",
    "Gestione rilievo": commentsToText(Array.isArray(r.comments) ? r.comments : []),
    "Immagine BCF": r.snapshotDataUrl ? "Presente" : "",
    Stato: translateStatus(r.stato),
  }));
}

function toChartExportRows(rows: any[]) {
  return rows.map((r: any) => ({
    Voce: r.label,
    Totale: r.value,
    NC: r.nc ?? "",
    OSS: r.oss ?? "",
  }));
}

function ExportButton({ children, onClick }: any) {
  return (
    <button
      onClick={onClick}
      style={{
        background: "#0f172a",
        color: "white",
        border: "none",
        borderRadius: 10,
        padding: "8px 12px",
        cursor: "pointer",
        fontWeight: 700,
        fontSize: 12,
      }}
    >
      {children}
    </button>
  );
}

function Card({ children, onClick, active = false }: any) {
  return (
    <div
      onClick={onClick}
      style={{
        background: active ? "#e0f2fe" : "white",
        borderRadius: 16,
        padding: 16,
        border: active ? "2px solid #0284c7" : "1px solid #e2e8f0",
        cursor: onClick ? "pointer" : "default",
        boxShadow: "0 6px 18px rgba(15,23,42,.06)",
      }}
    >
      {children}
    </div>
  );
}

function KPI({ title, value, subtitle, onClick, active, colorValue }: any) {
  return (
    <Card onClick={onClick} active={active}>
      <div style={{ fontSize: 13, color: "#64748b" }}>{title}</div>
      <div style={{ fontSize: 34, fontWeight: 800, color: colorValue || "#0f172a" }}>
        {value}
      </div>
      {subtitle && <div style={{ fontSize: 12, color: "#64748b" }}>{subtitle}</div>}
    </Card>
  );
}

function BarList({ title, data, onClick, activeKey, onExport }: any) {
  const max = Math.max(...data.map((d: any) => d.value), 1);

  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
        <h3 style={{ marginTop: 0 }}>{title}</h3>
        {onExport && <ExportButton onClick={onExport}>Export Excel</ExportButton>}
      </div>

      {data.length === 0 && (
        <div style={{ color: "#64748b", fontSize: 13 }}>Nessun dato disponibile</div>
      )}

      {data.map((d: any) => {
        const itemKey = d.key || d.label;

        return (
        <div
          key={itemKey}
          onClick={() => onClick(itemKey, d.label)}
          style={{
            marginBottom: 12,
            cursor: "pointer",
            background: activeKey === itemKey ? "#e0f2fe" : "transparent",
            borderRadius: 10,
            padding: 6,
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
            <span>{d.label}</span>
            <b>{d.value}</b>
          </div>

          {(d.nc !== undefined || d.oss !== undefined) && (
            <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>
              NC: {d.nc || 0} · OSS: {d.oss || 0}
            </div>
          )}

          <div
            style={{
              height: 12,
              background: "#e2e8f0",
              borderRadius: 99,
              overflow: "hidden",
              marginTop: 4,
            }}
          >
            <div
              style={{
                width: `${(d.value / max) * 100}%`,
                height: 12,
                background: "#0f172a",
              }}
            />
          </div>
        </div>
        );
      })}
    </Card>
  );
}

function ImportSummary({ importedFiles }: any) {
  if (!importedFiles?.length) return null;

  return (
    <Card>
      <h3 style={{ marginTop: 0 }}>File importati</h3>
      <div style={{ display: "grid", gap: 8 }}>
        {importedFiles.map((f: any, i: number) => (
          <div key={`${f.fileName}-${i}`} style={{ fontSize: 13 }}>
            <b>{f.fileName}</b>{" "}
            <span style={{ color: "#64748b" }}>
              {f.type === "xlsx" && `- Excel letto: ${f.rows || 0} righe`}
              {(f.type === "bcf" || f.type === "bcfzip" || f.type === "zip") &&
                `- BCF letto: ${f.markupCount || 0} topic, ${f.comments || 0} commenti${f.snapshots ? `, ${f.snapshots} immagini` : ""}${f.docx ? `, ${f.docx} schede Word` : ""}`}
              {f.type === "docx" &&
                `- Scheda Word letta: ${f.rilievi || 0} rilievi, ${f.elaboratiSenzaRilievi || 0} elaborati senza rilievi${f.disciplina ? ` - ${f.disciplina}` : ""}`}
            </span>
          </div>
        ))}
      </div>
    </Card>
  );
}


function isKnownInspectorAuthor(author: any) {
  const normalized = normalizeCommentAuthor(author || "").toLowerCase();
  if (!normalized) return false;

  return SIGNATURE_DATABASE.some((entry) => {
    const entryName = normalizeCommentAuthor(entry.name || "").toLowerCase();
    return entry.role === "Ispettore Tecnico" && entryName && normalized.includes(entryName);
  });
}

function isProjectComment(comment: any) {
  const role = String(comment?.role || "").toUpperCase();
  const author = comment?.author || "";

  if (role === "PRG") return true;
  if (role === "ISP") return false;

  // Se l'autore non è un ispettore noto, lo considero progettista.
  return !isKnownInspectorAuthor(author);
}

function hasProjectComment(row: any) {
  const comments = Array.isArray(row?.comments) ? row.comments : [];
  return Boolean(row?.hasPrgComment) || comments.some((comment: any) => isProjectComment(comment));
}

function getTodoQualityIssues(row: any) {
  const tipo = cleanPdfText(row?.tipo);
  const tipoUpper = tipo.toUpperCase();
  const disciplina = cleanPdfText(getDisciplinaDisplay(row));
  const elaborato = cleanPdfText(getElaboratoKey(row));
  const descrizione = cleanPdfText(row?.descrizione);
  const stato = cleanPdfText(row?.stato || translateStatus(row?.stato));

  const issues: Record<string, string> = {};

  if (!elaborato || elaborato === "Elaborato non identificato") {
    issues.elaborato = "Elaborato mancante";
  }

  if (!disciplina || disciplina === "Non assegnata" || disciplina.toLowerCase() === "null null") {
    issues.disciplina = "Disciplina mancante o non riconosciuta";
  }

  if (!tipo || tipoUpper === "ESITO MANCANTE" || tipoUpper === "RILIEVO MANCANTE") {
    issues.tipo = "Tipologia NC/OSS mancante";
  }

  if ((tipoUpper === "NC" || tipoUpper === "OSS" || tipoUpper === "DA NC A OSS") && !descrizione) {
    issues.descrizione = "Descrizione mancante";
  }

  if (tipoUpper === "NESSUN RILIEVO" && stato && stato !== "Chiusa") {
    issues.stato = "Stato incoerente con Nessun rilievo";
  }

  if ((tipoUpper === "NC" || tipoUpper === "OSS" || tipoUpper === "DA NC A OSS") && stato === "Chiusa" && !hasProjectComment(row)) {
    issues.gestione = "Rilievo chiuso senza commento progettista";
  }

  if (
    elaborato.toUpperCase() === "RILIEVO GENERALE" ||
    elaborato.toUpperCase() === "RILIEVI GENERALI" ||
    elaborato.toUpperCase() === "OSSERVAZIONI GENERALI"
  ) {
    issues.elaborato = "Rilievo generale: associare il rilievo a uno o più elaborati specifici";
  }

  return issues;
}

function hasTodoQualityIssues(row: any) {
  return Object.keys(getTodoQualityIssues(row)).length > 0;
}

function anomalyCellStyle(baseStyle: any, issue?: string) {
  if (!issue) return baseStyle;

  return {
    ...baseStyle,
    background: "#fee2e2",
    border: "2px solid #ef4444",
    color: "#7f1d1d",
    fontWeight: 700,
  };
}

function CommentList({ comments, emptyText = "" }: any) {
  if (!comments || comments.length === 0) {
    return <span style={{ color: "#94a3b8" }}>{emptyText}</span>;
  }

  return (
    <div style={{ minWidth: 320 }}>
      {comments.map((c: any, idx: number) => (
        <div
          key={`${c.date || "data"}-${c.author || "autore"}-${idx}`}
          style={{
            marginBottom: 10,
            paddingBottom: 10,
            borderBottom: idx < comments.length - 1 ? "1px solid #e2e8f0" : "none",
            whiteSpace: "pre-wrap",
          }}
        >
          <div style={{ fontWeight: 700 }}>
            {c.role ? `[${c.role}] ` : ""}
            {c.author || "Autore non indicato"}
          </div>
          {c.date && <div style={{ color: "#64748b", fontSize: 12 }}>{c.date}</div>}
          <div style={{ marginTop: 4 }}>{c.comment || ""}</div>
        </div>
      ))}
    </div>
  );
}


function ElaboratiPerDisciplinaPanel({ data, activeKey, onClick, onExport }: any) {
  const grouped = (data || []).reduce((acc: Record<string, any[]>, item: any) => {
    const key = item.disciplina || "Non assegnata";
    acc[key] = acc[key] || [];
    acc[key].push(item);
    return acc;
  }, {});

  const discipline = Object.keys(grouped).sort((a, b) => a.localeCompare(b, "it"));

  function renderItems(items: any[], tipoLabel: string) {
    if (!items || items.length === 0) return null;

    return (
      <div style={{ marginTop: 8 }}>
        <div style={{ fontSize: 12, fontWeight: 800, color: "#475569", marginBottom: 6 }}>
          {tipoLabel} ({items.length})
        </div>
        <div style={{ display: "grid", gap: 6 }}>
          {items.map((item: any) => {
            const isActive = activeKey === item.elaboratoKey;
            return (
              <div
                key={item.key}
                onClick={() => onClick(item.elaboratoKey, item.elaborato)}
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 70px 70px 120px",
                  gap: 8,
                  alignItems: "center",
                  cursor: "pointer",
                  padding: 7,
                  borderRadius: 8,
                  background: isActive ? "#e0f2fe" : "white",
                  border: isActive ? "2px solid #0284c7" : "1px solid #e2e8f0",
                  fontSize: 12,
                }}
                title={item.elaborato}
              >
                <div style={{ fontWeight: 700, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {item.elaborato}
                </div>
                <div>NC: <b>{item.nc}</b></div>
                <div>OSS: <b>{item.oss}</b></div>
                <div style={{ color: item.nc === 0 && item.oss === 0 ? "#15803d" : "#0f172a" }}>
                  {item.esito}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
        <h3 style={{ marginTop: 0 }}>Elaborati / Modelli per disciplina</h3>
        <ExportButton onClick={onExport}>Export Excel</ExportButton>
      </div>

      {discipline.length === 0 && (
        <div style={{ color: "#64748b", fontSize: 13 }}>Nessun dato disponibile</div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(360px, 1fr))", gap: 14 }}>
        {discipline.map((disciplina) => {
          const items = grouped[disciplina] || [];
          const modelli = items.filter((item: any) => item.tipoOggetto === "modello");
          const elaborati = items.filter((item: any) => item.tipoOggetto !== "modello");

          return (
            <div
              key={disciplina}
              style={{
                border: "1px solid #e2e8f0",
                borderRadius: 12,
                padding: 12,
                background: "#f8fafc",
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center", marginBottom: 8 }}>
                <div style={{ fontWeight: 800 }}>{disciplina}</div>
                <div style={{ fontSize: 12, color: "#475569", fontWeight: 700 }}>
                  {elaborati.length} elaborati
                </div>
              </div>

              {renderItems(elaborati, "Elaborati")}
              {renderItems(modelli, "Modelli 3D")}
            </div>
          );
        })}
      </div>
    </Card>
  );
}

function DetailPanel({ rows, title, onReset }: any) {
  const [pdfHeader, setPdfHeader] = useState<PdfHeaderData>({});
  const [showOnlyAnomalies, setShowOnlyAnomalies] = useState(false);

  if (!title) return null;

  const visibleRows = showOnlyAnomalies ? rows.filter((r: any) => hasTodoQualityIssues(r)) : rows;

  function updatePdfHeader(field: keyof PdfHeaderData, value: string) {
    setPdfHeader((prev) => ({ ...prev, [field]: value }));
  }

  function selectResponsabileTecnico(key: string) {
    const entry = getSignatureEntryByKey(key);
    setPdfHeader((prev) => ({
      ...prev,
      responsabileTecnicoKey: key,
      responsabileTecnico: entry?.name || "",
      firmaResponsabileImage: entry?.signature || "",
    }));
  }

  function toggleIspettoreTecnico(key: string) {
    setPdfHeader((prev) => {
      const currentKeys = prev.ispettoreKeys || [];
      const nextKeys = currentKeys.includes(key)
        ? currentKeys.filter((item) => item !== key)
        : [...currentKeys, key];
      const entries = nextKeys
        .map((item) => getSignatureEntryByKey(item))
        .filter(Boolean) as SignatureDatabaseEntry[];

      return {
        ...prev,
        ispettoreKeys: nextKeys,
        ispettore: entries.map((entry) => entry.name).join(" / "),
        firmaIspettoreImages: entries.map((entry) => entry.signature).filter(Boolean),
      };
    });
  }


  function updateFirmaResponsabileImage(file?: File) {
    if (!file) {
      setPdfHeader((prev) => ({ ...prev, firmaResponsabileImage: "" }));
      return;
    }

    if (!isAllowedSignatureImage(file)) {
      alert("Caricare una firma in formato .png, .jpg o .jpeg");
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      setPdfHeader((prev) => ({
        ...prev,
        firmaResponsabileImage: typeof reader.result === "string" ? reader.result : "",
      }));
    };
    reader.readAsDataURL(file);
  }

  function updateFirmaIspettoreImages(files?: FileList | null) {
    const selectedFiles = Array.from(files || []);

    if (selectedFiles.some((file) => !isAllowedSignatureImage(file))) {
      alert("Caricare solo firme in formato .png, .jpg o .jpeg");
      return;
    }

    if (selectedFiles.length === 0) {
      setPdfHeader((prev) => ({ ...prev, firmaIspettoreImages: [] }));
      return;
    }

    Promise.all(
      selectedFiles.map(
        (file) =>
          new Promise<string>((resolve) => {
            const reader = new FileReader();
            reader.onload = () => resolve(typeof reader.result === "string" ? reader.result : "");
            reader.readAsDataURL(file);
          })
      )
    ).then((images) => {
      setPdfHeader((prev) => ({
        ...prev,
        firmaIspettoreImages: images.filter(Boolean),
      }));
    });
  }

  const inputStyle = {
    width: "100%",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    padding: "8px 10px",
    fontSize: 12,
    background: "white",
  };

  const labelStyle = {
    display: "block",
    fontSize: 11,
    color: "#475569",
    fontWeight: 700,
    marginBottom: 4,
  };

  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
        <h2 style={{ marginTop: 0 }}>Dettaglio selezione: {title}</h2>
        <div style={{ display: "flex", gap: 10 }}>
          <ExportButton
            onClick={() =>
              exportExcel(
                "Dettaglio_selezione",
                toDashboardExportRows(visibleRows)
              )
            }
          >
            Export Excel
          </ExportButton>
          <ExportButton onClick={() => exportDetailPdf(visibleRows, title, pdfHeader)}>
            Export PDF
          </ExportButton>
          <ExportButton onClick={() => exportDetailPdf(rows, title, { templateMode: true } as any)}>
            Export PDF Template Qualità
          </ExportButton>
          <label
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              fontSize: 12,
              color: "#475569",
              border: "1px solid #cbd5e1",
              background: showOnlyAnomalies ? "#fee2e2" : "white",
              borderRadius: 10,
              padding: "8px 10px",
              height: 38,
            }}
          >
            <input
              type="checkbox"
              checked={showOnlyAnomalies}
              onChange={(e) => setShowOnlyAnomalies(e.target.checked)}
            />
            Solo anomalie
          </label>
          <button
            onClick={onReset}
            style={{
              border: "1px solid #cbd5e1",
              background: "white",
              borderRadius: 10,
              padding: "8px 12px",
              cursor: "pointer",
              height: 38,
            }}
          >
            Reset selezione
          </button>
        </div>
      </div>

      <div
        style={{
          marginTop: 12,
          marginBottom: 14,
          padding: 12,
          border: "1px solid #e2e8f0",
          borderRadius: 12,
          background: "#f8fafc",
        }}
      >
        <div style={{ fontWeight: 800, marginBottom: 8, color: "#0f172a" }}>
          Dati testata PDF - compilazione opzionale
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
          <div>
            <label style={labelStyle}>Committente / Stazione Appaltante</label>
            <input
              style={inputStyle}
              value={pdfHeader.committente || ""}
              onChange={(e) => updatePdfHeader("committente", e.target.value)}
              placeholder="es. RPR - Risorse per Roma S.p.A."
            />
          </div>

          <div>
            <label style={labelStyle}>Nota di Ricezione Elaborati e data</label>
            <input
              style={inputStyle}
              value={pdfHeader.notaRicezione || ""}
              onChange={(e) => updatePdfHeader("notaRicezione", e.target.value)}
              placeholder="es. 25063AR-PE1-RE-0001"
            />
          </div>
          <div>
            <label style={labelStyle}>Codice Scheda</label>
            <input
              style={inputStyle}
              value={pdfHeader.codiceScheda || ""}
              onChange={(e) => updatePdfHeader("codiceScheda", e.target.value)}
              placeholder="es. SP-001 o Allegato 1"
            />
          </div>
          <div>
            <label style={labelStyle}>Oggetto del Contratto / Accordo Quadro</label>
            <textarea
              style={{
                ...inputStyle,
                minHeight: 90,
                resize: "vertical",
                lineHeight: "1.35",
              }}
              maxLength={300}
              value={pdfHeader.oggetto || ""}
              onChange={(e) => updatePdfHeader("oggetto", e.target.value.slice(0, 300))}
              placeholder="Descrizione sintetica della commessa fino a 300 caratteri"
            />
            <div style={{ fontSize: 11, color: "#64748b", marginTop: 4, textAlign: "right" }}>
              {(pdfHeader.oggetto || "").length}/300
            </div>
          </div>
          <div>
            <label style={labelStyle}>Nome Responsabile tecnico</label>
            <select
              style={inputStyle}
              value={pdfHeader.responsabileTecnicoKey || ""}
              onChange={(e) => selectResponsabileTecnico(e.target.value)}
            >
              <option value="">Seleziona Responsabile tecnico</option>
              {getResponsabileOptions().map((entry) => {
                const key = signatureKey(entry);
                return (
                  <option key={key} value={key}>
                    {entry.name} - {entry.discipline}
                  </option>
                );
              })}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Nome ispettore</label>
            <input
              style={inputStyle}
              value={pdfHeader.ispettore || ""}
              readOnly
              placeholder="Seleziona uno o più ispettori dall'elenco"
            />
          </div>
          <div>
            <label style={labelStyle}>Data emissione</label>
            <input
              style={inputStyle}
              value={pdfHeader.dataEmissione || ""}
              onChange={(e) => updatePdfHeader("dataEmissione", e.target.value)}
              placeholder="Data emissione"
            />
          </div>
          <div style={{ gridColumn: "1 / -1" }}>
            <label style={labelStyle}>Ispettori tecnici e firme</label>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
                gap: 8,
                border: "1px solid #cbd5e1",
                borderRadius: 8,
                padding: 10,
                background: "white",
              }}
            >
              {getIspettoreOptions().map((entry) => {
                const key = signatureKey(entry);
                const checked = Boolean(pdfHeader.ispettoreKeys?.includes(key));

                return (
                  <label
                    key={key}
                    style={{
                      display: "grid",
                      gridTemplateColumns: "18px 1fr 120px",
                      alignItems: "center",
                      gap: 8,
                      padding: 6,
                      border: checked ? "1px solid #0ea5e9" : "1px solid #e2e8f0",
                      borderRadius: 8,
                      background: checked ? "#e0f2fe" : "#f8fafc",
                      cursor: "pointer",
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={() => toggleIspettoreTecnico(key)}
                    />
                    <span>
                      <strong>{entry.name}</strong>
                      <br />
                      <span style={{ color: "#64748b", fontSize: 11 }}>{entry.discipline}</span>
                    </span>
                    {checked && entry.signature ? (
                      <img src={entry.signature} alt={`Firma ${entry.name}`} style={{ maxWidth: 110, maxHeight: 34, objectFit: "contain" }} />
                    ) : (
                      <span style={{ color: "#94a3b8", fontSize: 11 }}>firma</span>
                    )}
                  </label>
                );
              })}
            </div>
          </div>
        </div>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#f1f5f9" }}>
              <th style={th}>Qualità</th>
              <th style={th}>N.</th>
              <th style={th}>ID_Rilievo</th>
              <th style={th}>Tipologia rilievo</th>
              <th style={th}>Disciplina</th>
              <th style={th}>Redattore</th>
              <th style={th}>Elaborato</th>
              <th style={th}>Descrizione</th>
              <th style={th}>Gestione rilievo</th>
              <th style={th}>Stato</th>
            </tr>
          </thead>

          <tbody>
            {visibleRows.map((r: any, i: number) => {
              const allComments = Array.isArray(r.comments) ? r.comments : [];
              const issues = getTodoQualityIssues(r);
              const issueList = Object.values(issues).join("; ");
              const rowHasIssues = issueList.length > 0;

              return (
                <tr
                  key={`${r.id}-${i}`}
                  style={{
                    background: rowHasIssues ? "#fff7ed" : "transparent",
                  }}
                >
                  <td
                    style={{
                      ...td,
                      textAlign: "center",
                      fontWeight: 800,
                      color: rowHasIssues ? "#b45309" : "#15803d",
                    }}
                    title={rowHasIssues ? issueList : "Nessuna anomalia"}
                  >
                    {rowHasIssues ? "⚠" : "✓"}
                  </td>
                  <td style={{ ...td, textAlign: "center", fontWeight: 700 }}>{i + 1}</td>
                  <td style={td}>{r.id}</td>
                  <td style={anomalyCellStyle(td, issues.tipo)} title={issues.tipo || ""}>{r.tipo}</td>
                  <td style={anomalyCellStyle(td, issues.disciplina)} title={issues.disciplina || ""}>{getDisciplinaDisplay(r)}</td>
                  <td style={td}>{getRedattoreFromRow(r)}</td>
                  <td style={anomalyCellStyle(td, issues.elaborato)} title={issues.elaborato || ""}>{getElaboratoKey(r)}</td>
                  <td style={anomalyCellStyle(td, issues.descrizione)} title={issues.descrizione || ""}>{r.descrizione}</td>
                  <td style={anomalyCellStyle(td, issues.gestione)} title={issues.gestione || ""}>
                    <CommentList comments={allComments} emptyText="Nessun commento" />
                  </td>
                  <td style={anomalyCellStyle(td, issues.stato)} title={issues.stato || ""}>{translateStatus(r.stato)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

const th = {
  border: "1px solid #e2e8f0",
  padding: 8,
  textAlign: "left" as const,
};

const td = {
  border: "1px solid #e2e8f0",
  padding: 8,
  verticalAlign: "top" as const,
};

const defaultDashboardModules = [
  {
    title: "Nota di Ricezione Elaborati",
    subtitle: "Modulo operativo attivo",
    url: "https://verifica-elaborati-production.up.railway.app",
    active: true,
    visible: true,
    external: true,
    sort_order: 1,
  },
];

export default function AppProgettiUpload() {
  React.useEffect(() => {
    const isAuthenticated = localStorage.getItem("nexcommon_verify_auth");

    if (isAuthenticated !== "true") {
      window.location.href = "/login";
    }
  }, []);

  const [files, setFiles] = useState<File[]>([]);
  const [rows, setRows] = useState<any[]>([]);
  const [selection, setSelection] = useState<any>(null);
  const [importedFiles, setImportedFiles] = useState<any[]>([]);
  const [dashboardModules, setDashboardModules] = useState<any[]>(defaultDashboardModules);
  const [error, setError] = useState("");

  React.useEffect(() => {
    async function loadDashboardModules() {
      const { data, error } = await supabase
        .from("dashboard_modules")
        .select("*")
        .eq("visible", true)
        .order("sort_order", { ascending: true });

      if (error) {
        console.error("Errore caricamento moduli dashboard:", error);
        return;
      }

      setDashboardModules(data && data.length > 0 ? data : defaultDashboardModules);
    }

    loadDashboardModules();
  }, []);

  async function generaDashboard() {
    setError("");

    const fd = new FormData();
    files.forEach((f) => fd.append("files", f));

    const res = await fetch("/api/parse", {
      method: "POST",
      body: fd,
    });

    const data = await res.json();

    if (!res.ok || data.ok === false) {
      setError(data.error || "Errore durante la lettura dei file");
      setRows([]);
      setImportedFiles([]);
      setSelection(null);
      return;
    }

    setRows(data.rows || []);
    setImportedFiles(data.importedFiles || []);
    setSelection(null);
  }

  const enrichedRows = useMemo(() => {
    return rows.map((r) => ({
      ...r,
      tipologiaNcOss: r.tipologiaNcOss || r.tipologiaDocumento || "",
      tipologia: r.tipologiaNcOss || r.tipologiaDocumento || "",
      elaboratoKey: getElaboratoNormalizedKey(r),
      elaboratoDisplay: getElaboratoDisplay(r),
      stato: translateStatus(r.stato),
    }));
  }, [rows]);

  const elaboratiTot = new Set(enrichedRows.map((r) => r.elaboratoKey).filter(Boolean)).size;
  const elaboratiNC = new Set(enrichedRows.filter((r) => r.tipo === "NC").map((r) => r.elaboratoKey).filter(Boolean)).size;
  const elaboratiOSS = new Set(enrichedRows.filter((r) => r.tipo === "OSS").map((r) => r.elaboratoKey).filter(Boolean)).size;
  const elaboratiOK = new Set(enrichedRows.filter((r) => r.tipo === "Nessun rilievo").map((r) => r.elaboratoKey).filter(Boolean)).size;

  const totaleNC = enrichedRows.filter((r) => r.tipo === "NC").length;
  const totaleOSS = enrichedRows.filter((r) => r.tipo === "OSS").length;
  const totaleNCOSS = totaleNC + totaleOSS;

  const rilieviNCOSS = enrichedRows.filter(
    (r) => r.tipo === "NC" || r.tipo === "OSS" || r.tipo === "Da NC a OSS"
  );

  const daVerificareISP = enrichedRows.filter((r) => r.chiDeveAgire === "ISP").length;
  const daRisponderePRG = enrichedRows.filter((r) => r.chiDeveAgire === "PRG").length;



  const discipline: any = {};
  const esiti: any = {};
  const rilieviPerElaborato: any = {};

  enrichedRows.forEach((r) => {
    const d = getDisciplinaDisplay(r);
    discipline[d] = (discipline[d] || 0) + 1;

    const e = r.tipo || "Rilievo mancante";
    esiti[e] = (esiti[e] || 0) + 1;

    const elaboratoKey = r.elaboratoKey || "Elaborato non identificato";
    const elaboratoLabel = r.elaboratoDisplay || getElaboratoDisplay(r);

    if (!rilieviPerElaborato[elaboratoKey]) {
      rilieviPerElaborato[elaboratoKey] = {
        key: elaboratoKey,
        label: elaboratoLabel,
        value: 0,
        nc: 0,
        oss: 0,
      };
    }

    if (r.tipo === "NC") {
      rilieviPerElaborato[elaboratoKey].value += 1;
      rilieviPerElaborato[elaboratoKey].nc += 1;
    }

    if (r.tipo === "OSS") {
      rilieviPerElaborato[elaboratoKey].value += 1;
      rilieviPerElaborato[elaboratoKey].oss += 1;
    }
  });

  const disciplineData = Object.entries(discipline)
    .map(([label, value]) => ({ label, value: Number(value) }))
    .sort((a: any, b: any) => b.value - a.value);

  const esitiData = Object.entries(esiti)
    .map(([label, value]) => ({ label, value: Number(value) }))
    .sort((a: any, b: any) => b.value - a.value);

  const rilieviPerElaboratoData = Object.values(rilieviPerElaborato)
    .filter((d: any) => d.value > 0)
    .sort((a: any, b: any) => b.value - a.value);

  const coverageRows = useMemo(() => {
    const grouped: Record<string, any> = {};

    enrichedRows.forEach((r: any) => {
      const key = r.elaboratoKey || normalizeElaboratoCode(getElaboratoKey(r));
      if (!key) return;

      if (!grouped[key]) {
        grouped[key] = {
          key,
          elaborato: r.elaboratoDisplay || getElaboratoDisplay(r),
          disciplina: getDisciplinaDisplay(r),
          tipoOggetto: getOggettoVerificaTipo(r),
          nc: 0,
          oss: 0,
          nessunRilievo: 0,
          esitoMancante: 0,
          commentato: false,
        };
      }

      const tipo = cleanPdfText(r.tipo).toUpperCase();
      if (tipo === "NC") grouped[key].nc += 1;
      if (tipo === "OSS" || tipo === "DA NC A OSS") grouped[key].oss += 1;
      if (tipo === "NESSUN RILIEVO") grouped[key].nessunRilievo += 1;
      if (!tipo || tipo === "ESITO MANCANTE" || tipo === "RILIEVO MANCANTE") grouped[key].esitoMancante += 1;

      if (getOggettoVerificaTipo(r) === "modello") {
        grouped[key].tipoOggetto = "modello";
        grouped[key].disciplina = "BIM";
      }

      if (tipo === "NC" || tipo === "OSS" || tipo === "DA NC A OSS" || tipo === "NESSUN RILIEVO") {
        grouped[key].commentato = true;
      }
    });

    return Object.values(grouped).sort((a: any, b: any) => a.elaborato.localeCompare(b.elaborato, "it"));
  }, [enrichedRows]);

  const elaboratiCommentati = coverageRows.filter((r: any) => r.commentato).length;
  const elaboratiNonCommentati = coverageRows.filter((r: any) => !r.commentato || r.esitoMancante > 0).length;
  const percentualeCopertura = coverageRows.length > 0
    ? Math.round((elaboratiCommentati / coverageRows.length) * 100)
    : 0;

  const coverageExportRows = coverageRows.map((r: any) => ({
    Elaborato: r.elaborato,
    Disciplina: r.disciplina,
    NC: r.nc,
    OSS: r.oss,
    "Nessun rilievo": r.nessunRilievo,
    "Esito mancante": r.esitoMancante,
    Stato: r.commentato && r.esitoMancante === 0 ? "Commentato" : "Da verificare",
  }));

  const elaboratiPerDisciplinaRows = useMemo(() => {
    return coverageRows.map((r: any) => {
      const tipoOggetto = r.tipoOggetto === "modello" ? "modello" : "elaborato";

      return {
        key: `${r.disciplina || "Non assegnata"}__${tipoOggetto}__${r.key}`,
        disciplina: r.disciplina || "Non assegnata",
        tipoOggetto,
        elaborato: r.elaborato,
        nc: r.nc,
        oss: r.oss,
        nessunRilievo: r.nessunRilievo,
        esito: r.nc === 0 && r.oss === 0 ? "Nessuna NC/OSS" : `NC: ${r.nc} - OSS: ${r.oss}`,
        value: r.nc + r.oss + r.nessunRilievo,
        elaboratoKey: r.key,
      };
    });
  }, [coverageRows]);

  const elaboratiPerDisciplinaExportRows = elaboratiPerDisciplinaRows.map((r: any) => ({
    Disciplina: r.disciplina,
    Tipo: r.tipoOggetto === "modello" ? "Modello 3D IFC" : "Elaborato",
    Nome: r.elaborato,
    NC: r.nc,
    OSS: r.oss,
    "Nessun rilievo": r.nessunRilievo,
    Esito: r.esito,
  }));

  const todoQualityRows = useMemo(() => {
    return enrichedRows.map((r: any) => {
      const issuesMap = getTodoQualityIssues(r);
      const issues = Object.values(issuesMap);

      return {
        id: r.id || "",
        elaborato: cleanPdfText(getElaboratoKey(r)) || "Elaborato non identificato",
        disciplina: cleanPdfText(getDisciplinaDisplay(r)),
        tipo: cleanPdfText(r.tipo) || "Mancante",
        descrizionePresente: cleanPdfText(r.descrizione) ? "Sì" : "No",
        stato: cleanPdfText(r.stato),
        esito: issues.length === 0 ? "OK" : "Da verificare",
        anomalie: issues.join("; "),
      };
    }).filter((item: any) => item.esito !== "OK");
  }, [enrichedRows]);

  const filteredRows = useMemo(() => {
    if (!selection) return [];

    if (selection.type === "kpi") {
      if (selection.value === "totali") return enrichedRows;
      if (selection.value === "nc") return enrichedRows.filter((r) => r.tipo === "NC");
      if (selection.value === "oss") return enrichedRows.filter((r) => r.tipo === "OSS");
      if (selection.value === "nessun") return enrichedRows.filter((r) => r.tipo === "Nessun rilievo");
      if (selection.value === "risoluzione-prg") return rilieviNCOSS.filter((r) => r.hasPrgComment);
      if (selection.value === "da-verificare-isp") return enrichedRows.filter((r) => r.chiDeveAgire === "ISP");
      if (selection.value === "da-rispondere-prg") return enrichedRows.filter((r) => r.chiDeveAgire === "PRG");
      if (selection.value === "copertura-mancanti") {
        const missingKeys = new Set(coverageRows.filter((r: any) => !r.commentato || r.esitoMancante > 0).map((r: any) => r.key));
        return enrichedRows.filter((r) => missingKeys.has(r.elaboratoKey));
      }
      if (selection.value === "qualita-todo") {
        const issueIds = new Set(todoQualityRows.map((r: any) => r.id).filter(Boolean));
        return enrichedRows.filter((r) => issueIds.has(r.id));
      }
    }

    if (selection.type === "tipo") return enrichedRows.filter((r) => r.tipo === selection.value);
    if (selection.type === "disciplina") return enrichedRows.filter((r) => getDisciplinaDisplay(r) === selection.value);
    if (selection.type === "elaborato") return enrichedRows.filter((r) => r.elaboratoKey === selection.value);
    if (selection.type === "elaborato-disciplina") return enrichedRows.filter((r) => r.elaboratoKey === selection.value);

    return [];
  }, [enrichedRows, selection, rilieviNCOSS, coverageRows, todoQualityRows]);

  const selectionTitle = selection
    ? `${selection.label}: ${selection.valueLabel || selection.value}`
    : "";

  return (
    <main
      style={{
        padding: 30,
        background: "#f1f5f9",
        minHeight: "100vh",
        fontFamily: "Arial, sans-serif",
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          gap: 20,
          alignItems: "flex-start",
        }}
      >
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <img src="/logo_nexcommon.png" alt="Nexcommon" style={{ height: 34, objectFit: "contain" }} />
            <div style={{ fontSize: 13, color: "#64748b" }}>
              Piattaforma creata da Nexcommon S.r.l.
            </div>
          </div>

          <div style={{ marginTop: 22, display: "flex", alignItems: "center", gap: 18 }}>
            <img
              src="/logo_its.png"
              alt="ITS Controlli Tecnici S.p.A."
              style={{
                height: 58,
                objectFit: "contain",
                background: "#0f172a",
                padding: 8,
                borderRadius: 8,
              }}
            />

            <div>
              <h1 style={{ margin: 0, fontSize: 30 }}>
                ITS Controlli Tecnici S.p.A.
              </h1>
              <div style={{ color: "#64748b", fontSize: 14 }}>
                Dashboard verifiche elaborati / NC / OSS
              </div>
            </div>
          </div>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <button
            onClick={() => {
              localStorage.removeItem("nexcommon_verify_auth");
              window.location.href = "/login";
            }}
            style={{
              alignSelf: "flex-end",
              background: "#dc2626",
              color: "white",
              border: "none",
              borderRadius: 10,
              padding: "10px 14px",
              fontWeight: 700,
              cursor: "pointer",
            }}
          >
            Logout
          </button>

          <Card>
            <input
              type="file"
              multiple
              accept=".xlsx,.xls,.bcf,.bcfzip,.zip,.docx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={(e) => setFiles(Array.from(e.target.files || []))}
            />
            <button
              onClick={generaDashboard}
              disabled={!files.length}
              style={{
                marginTop: 10,
                width: "100%",
                padding: 10,
                background: "#0f172a",
                color: "white",
                borderRadius: 10,
                border: "none",
                cursor: files.length ? "pointer" : "not-allowed",
              }}
            >
              Analizza ToDo / BCF / BCFZIP / SCHEDE WORD
            </button>

            {files.length > 0 && (
              <div style={{ marginTop: 10, color: "#64748b", fontSize: 12 }}>
                File selezionati: {files.map((f) => f.name).join(", ")}
              </div>
            )}
          </Card>
        </div>
      </div>

      {error && (
        <div
          style={{
            marginTop: 16,
            background: "#fee2e2",
            color: "#991b1b",
            border: "1px solid #fecaca",
            borderRadius: 12,
            padding: 12,
          }}
        >
          {error}
        </div>
      )}

      <div style={{ marginTop: 16 }}>
        <ImportSummary importedFiles={importedFiles} />
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
          gap: 12,
          marginTop: 24,
          marginBottom: 24,
        }}
      >
        {dashboardModules
          .filter((m) => m.visible)
          .filter((m) => m.code !== "dashboard-nc-oss" && m.url !== "/dashboard-pm")
          .map((m) => (
            <Card
              key={m.code || m.title}
              active={m.active}
              onClick={() => {
                if (!m.active || !m.url) return;

                if (m.external) {
                  window.open(m.url, "_blank", "noopener,noreferrer");
                  return;
                }

                window.location.href = m.url;
              }}
            >
              <b>{m.title}</b>
              <div style={{ marginTop: 6, color: "#64748b", fontSize: 13 }}>
                {m.subtitle}
              </div>

              {m.active ? (
                <div style={{ marginTop: 12, fontSize: 12, fontWeight: 700, color: "#0284c7" }}>
                  Apri modulo →
                </div>
              ) : (
                <div style={{ marginTop: 12, fontSize: 12, fontWeight: 700, color: "#94a3b8" }}>
                  Modulo in standby
                </div>
              )}
            </Card>
          ))}
      </div>

      {rows.length === 0 && (
        <Card>
          <div style={{ fontWeight: 700 }}>Nessun dato caricato</div>
          <div style={{ marginTop: 6, color: "#64748b", fontSize: 13 }}>
            Carica uno o più file ToDo .xlsx, .bcf, .bcfzip o schede ispettive .docx
            per generare la sintesi NC/OSS.
          </div>
        </Card>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 12, marginTop: 24, marginBottom: 12 }}>
        <KPI title="Elaborati totali" value={elaboratiTot} onClick={() => setSelection({ type: "kpi", value: "totali", label: "KPI", valueLabel: "Elaborati totali" })} />
        <KPI title="Elaborati con NC" value={elaboratiNC} onClick={() => setSelection({ type: "kpi", value: "nc", label: "KPI", valueLabel: "Elaborati con NC" })} />
        <KPI title="Elaborati con OSS" value={elaboratiOSS} onClick={() => setSelection({ type: "kpi", value: "oss", label: "KPI", valueLabel: "Elaborati con OSS" })} />
        <KPI title="Elaborati senza rilievi" value={elaboratiOK} onClick={() => setSelection({ type: "kpi", value: "nessun", label: "KPI", valueLabel: "Elaborati senza rilievi" })} />
        <KPI title="Totale NC" value={totaleNC} onClick={() => setSelection({ type: "tipo", value: "NC", label: "Rilievi" })} />
        <KPI title="Totale OSS" value={totaleOSS} onClick={() => setSelection({ type: "tipo", value: "OSS", label: "Rilievi" })} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 24 }}>
        <KPI title="Totale Rilievi" value={totaleNCOSS} onClick={() => setSelection({ type: "kpi", value: "totali", label: "KPI", valueLabel: "Tutti i rilievi" })} />
        <KPI title="In attesa di riscontro dell'ispettore" value={daVerificareISP} subtitle="Ultimo commento PRG" onClick={() => setSelection({ type: "kpi", value: "da-verificare-isp", label: "KPI", valueLabel: "Da verificare ISP" })} />
        <KPI title="In attesa di risposta del progettista" value={daRisponderePRG} subtitle="Nessun PRG o ultimo ISP" onClick={() => setSelection({ type: "kpi", value: "da-rispondere-prg", label: "KPI", valueLabel: "Da rispondere PRG" })} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
        <Card>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
            <h3 style={{ marginTop: 0 }}>Controllo copertura elaborati</h3>
            <ExportButton onClick={() => exportExcel("Controllo_copertura_elaborati", coverageExportRows)}>
              Export Excel
            </ExportButton>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10 }}>
            <KPI title="Da controllare" value={coverageRows.length} />
            <KPI title="Commentati" value={elaboratiCommentati} />
            <KPI title="Da verificare" value={elaboratiNonCommentati} onClick={() => setSelection({ type: "kpi", value: "copertura-mancanti", label: "Controllo", valueLabel: "Elaborati da verificare" })} colorValue={elaboratiNonCommentati > 0 ? "#b45309" : "#0f172a"} />
            <KPI title="Copertura" value={`${percentualeCopertura}%`} />
          </div>
          <div style={{ marginTop: 10, color: "#64748b", fontSize: 13 }}>
            Verifica se ogni elaborato ha un esito: NC, OSS, Da NC a OSS o Nessun rilievo.
          </div>
        </Card>

        <Card>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
            <h3 style={{ marginTop: 0 }}>Qualità compilazione ToDo</h3>
            <ExportButton onClick={() => exportExcel("Qualita_compilazione_ToDo", todoQualityRows)}>
              Export Excel
            </ExportButton>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
            <KPI title="ToDo controllati" value={enrichedRows.length} />
            <KPI title="Anomalie" value={todoQualityRows.length} onClick={() => setSelection({ type: "kpi", value: "qualita-todo", label: "Qualità ToDo", valueLabel: "Anomalie compilazione" })} colorValue={todoQualityRows.length > 0 ? "#b45309" : "#0f172a"} />
            <KPI title="Esito" value={todoQualityRows.length === 0 ? "OK" : "KO"} />
          </div>
          <div style={{ marginTop: 10, color: "#64748b", fontSize: 13 }}>
            Controlla campi obbligatori: elaborato, disciplina, tipologia, descrizione e coerenza stato.
          </div>
        </Card>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <BarList title="Rilievi per disciplina" data={disciplineData} activeKey={selection?.type === "disciplina" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "disciplina", value, label: "Disciplina" })} onExport={() => exportExcel("Rilievi_per_disciplina", toChartExportRows(disciplineData))} />
        <BarList title="Rilievi" data={esitiData} activeKey={selection?.type === "tipo" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "tipo", value, label: "Rilievi" })} onExport={() => exportExcel("Rilievi", toChartExportRows(esitiData))} />
      </div>

      <div style={{ marginTop: 16, marginBottom: 16 }}>
        <ElaboratiPerDisciplinaPanel
          data={elaboratiPerDisciplinaRows}
          activeKey={selection?.type === "elaborato-disciplina" ? selection.value : ""}
          onClick={(value: string, valueLabel: string) => setSelection({ type: "elaborato-disciplina", value, label: "Elaborato", valueLabel })}
          onExport={() => exportExcel("Elaborati_per_disciplina", elaboratiPerDisciplinaExportRows)}
        />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 16 }}>
        <BarList title="NC / OSS per elaborato" data={rilieviPerElaboratoData} activeKey={selection?.type === "elaborato" ? selection.value : ""} onClick={(value: string, valueLabel: string) => setSelection({ type: "elaborato", value, label: "Elaborato", valueLabel })} onExport={() => exportExcel("NC_OSS_per_elaborato", toChartExportRows(rilieviPerElaboratoData))} />
      </div>

      <div style={{ marginTop: 24 }}>
        <DetailPanel title={selectionTitle} rows={filteredRows} onReset={() => setSelection(null)} />
      </div>

      <div
        style={{
          marginTop: 40,
          padding: 20,
          background: "#0f172a",
          color: "white",
          borderRadius: 12,
          textAlign: "center",
          fontSize: 12,
        }}
      >
        <div style={{ fontWeight: 600 }}>Nexcommon S.r.l.</div>
        <div style={{ opacity: 0.8 }}>
          © {new Date().getFullYear()} – Tutti i diritti riservati
        </div>
        <div style={{ marginTop: 6, opacity: 0.7 }}>
          Piattaforma Quality Control per ITS Controlli Tecnici S.p.A.
        </div>
      </div>
    </main>
  );
}
