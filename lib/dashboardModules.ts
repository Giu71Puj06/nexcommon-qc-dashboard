export type DashboardModule = {
  code: string;
  title: string;
  description: string;
  href: string;
  visible: boolean;
  enabled: boolean;
};

export const dashboardModules: DashboardModule[] = [
  {
    code: "nota-ricezione",
    title: "Nota di Ricezione Elaborati",
    description: "Modulo operativo attivo",
    href: "/nota-ricezione",
    visible: true,
    enabled: true,
  },

  {
    code: "verifiche-preliminari",
    title: "Verifiche preliminari",
    description: "Modulo temporaneamente in standby",
    href: "/verifiche-preliminari",
    visible: true,
    enabled: false,
  },

  {
    code: "schede-ispettive",
    title: "Schede ispettive",
    description: "Modulo temporaneamente in standby",
    href: "/schede-ispettive",
    visible: true,
    enabled: false,
  },

  {
    code: "controllo-todo",
    title: "Controllo ToDo ispettori",
    description: "Modulo temporaneamente in standby",
    href: "/controllo-todo",
    visible: true,
    enabled: false,
  },

  {
    code: "rapporto-intermedio",
    title: "Rapporto intermedio",
    description: "Modulo temporaneamente in standby",
    href: "/rapporto-intermedio",
    visible: true,
    enabled: false,
  },

  {
    code: "rapporto-conclusivo",
    title: "Rapporto conclusivo",
    description: "Modulo temporaneamente in standby",
    href: "/rapporto-conclusivo",
    visible: true,
    enabled: false,
  },

  {
    code: "dashboard-nc-oss",
    title: "Dashboard NC / OSS",
    description: "Analisi ToDo, BCF e BCFZIP",
    href: "/dashboard-pm",
    visible: true,
    enabled: true,
  },
];
