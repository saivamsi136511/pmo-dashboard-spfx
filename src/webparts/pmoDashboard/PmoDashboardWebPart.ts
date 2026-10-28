import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import PmoDashboardRoot, { IPmoDashboardRootProps } from './components/PmoDashboardRoot';

// ─── Web Part Property Interface ──────────────────────────────────────────────

export interface IPmoDashboardWebPartProps {
  title: string;
}

// ─── Web Part Class ───────────────────────────────────────────────────────────

export default class PmoDashboardWebPart extends BaseClientSideWebPart<IPmoDashboardWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPmoDashboardRootProps> = React.createElement(
      PmoDashboardRoot,
      {
        context: this.context,
        title: this.properties.title || 'PMO Dashboard',
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // SPFx version — keep in sync with package-solution.json
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'PMO Dashboard Settings' },
          groups: [
            {
              groupName: 'Appearance',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title',
                  placeholder: 'e.g. Project Portfolio',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
