import * as React from 'react';
import { HashRouter, Route, Switch } from 'react-router-dom';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import DashboardView from './DashboardView';

export interface IPmoDashboardRootProps {
  context: WebPartContext;
  title: string;
}

// No QueryClientProvider needed — useQuery is a self-contained custom hook.
const PmoDashboardRoot: React.FC<IPmoDashboardRootProps> = ({ context, title }) => {
  return (
    <HashRouter>
      <Switch>
        <Route
          path="/"
          render={() => <DashboardView context={context} title={title} />}
        />
      </Switch>
    </HashRouter>
  );
};

export default PmoDashboardRoot;
