# Nexcommon Trimble QC

Starter kit per app Nexcommon collegata a Trimble ToDo/BCF tramite esportazione CSV.

## Moduli inclusi nell'MVP

- Anagrafica commesse
- Registro elaborati
- Import CSV Trimble ToDo/BCF
- Normalizzazione stati Trimble
- Lettura discipline dai gruppi
- Lettura esiti dai tag: NC, OSS, Nessun rilievo
- Tassonomia NC/OSS
- Calcolo completezza dati
- Semaforo PM
- Endpoint JSON per Power BI

## Setup locale

```bash
npm install
cp .env.example .env
# Inserire DATABASE_URL in .env
npx prisma migrate dev --name init
npm run seed
npm run dev
```

Aprire:

```text
http://localhost:3000
http://localhost:3000/dashboard
```

## Import CSV

Endpoint:

```text
POST /api/import-trimble
```

Form-data:

- file: CSV Trimble
- commessaCodice: codice commessa, esempio DEMO-001

File demo:

```text
sample-data/trimble-export-demo.csv
```

## Export Power BI

Endpoint:

```text
GET /api/export-powerbi
```

Power BI può leggere questo endpoint come sorgente Web JSON, oppure in fase successiva collegarsi direttamente a PostgreSQL.

## Deploy Railway

Railway supporta deploy da GitHub e app Next.js con PostgreSQL. Aggiungere un servizio PostgreSQL e collegare la variabile DATABASE_URL all'app.

Variabili richieste:

```text
DATABASE_URL
NEXT_PUBLIC_APP_NAME
```

`railway.json` esegue `npx prisma migrate deploy` prima dello start.
