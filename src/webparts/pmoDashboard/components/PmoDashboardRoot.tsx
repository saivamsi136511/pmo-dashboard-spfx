import * as React from 'react';
import { HashRouter, Route, Switch } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import DashboardView from './DashboardView';

// ─── Props ────────────────────────────────────────────────────────────────────

export interface IPmoDashboardRootProps {
  context: WebPartContext;
  title: string;
}

// ─── Singleton QueryClient ────────────────────────────────────────────────────
// Created once outside the component so it survives React re-renders and SPFx
// property-pane updates without losing the in-memory cache.

const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      // Cache data for 5 minutes before a background refetch is triggered.
      staleTime: 5 * 60 * 1000,
      // Keep data in cache for 10 minutes after all observers have unmounted.
      cacheTime: 10 * 60 * 1000,
      // Disable window-focus refetch to avoid noisy requests in SharePoint.
      refetchOnWindowFocus: false,
      // Retry once on failure before surfacing an error to the UI.
      retry: 1,
    },
  },
});

// ─── Root Component ───────────────────────────────────────────────────────────

const PmoDashboardRoot: React.FC<IPmoDashboardRootProps> = ({ context, title }) => {
  return (
    <QueryClientProvider client={queryClient}>
      {/*
        HashRouter is mandatory in SPFx: BrowserRouter rewrites the page URL
        which causes SharePoint to navigate away and throw a 404.
      */}
      <HashRouter>
        <Switch>
          <Route
            path="/"
            render={() => (
              <DashboardView context={context} title={title} />
            )}
          />
        </Switch>
      </HashRouter>
    </QueryClientProvider>
  );
};

export default PmoDashboardRoot;
