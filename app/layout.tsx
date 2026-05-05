import './globals.css';
import type { ReactNode } from 'react';
export const metadata = { title: 'Nexcommon Trimble QC', description: 'Hub verifiche elaborati Trimble ToDo/BCF' };
export default function RootLayout({ children }: { children: ReactNode }) {
  return <html lang="it"><body>{children}</body></html>;
}
