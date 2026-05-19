import Link from "next/link";
import AppProgettiUpload from "@/components/AppProgettiUpload";

export default function Home() {
  return (
    <>
      <AppProgettiUpload />

      <div style={{ padding: "24px" }}>
        <Link href="/dashboard-pm">
          <button>
            Apri Dashboard PM
          </button>
        </Link>
      </div>
    </>
  );
}
