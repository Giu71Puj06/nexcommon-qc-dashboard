export default function DashboardPMPage() {
  return (
    <div style={{ padding: "24px" }}>
      <h1>Dashboard PM</h1>

      <p>Modulo analisi BCF + XLSX</p>

      <input
        type="file"
        multiple
        accept=".bcf,.bcfzip,.xlsx"
      />
    </div>
  );
}
