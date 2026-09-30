import { useState } from 'react';
import { AppShell } from './components/AppShell/AppShell';
import { createPlatform } from './services/platform';
import { AppProvider, createAppServices, type AppServices } from './state/AppContext';

/** Root component. Tests pass their own services (mock platform, in-thread parser). */
export default function App({ services }: { services?: AppServices }) {
  const [appServices] = useState(() => services ?? createAppServices(createPlatform()));
  return (
    <AppProvider services={appServices}>
      <AppShell />
    </AppProvider>
  );
}
