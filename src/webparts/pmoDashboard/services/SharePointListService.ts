import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// ─── Data Model Interfaces ────────────────────────────────────────────────────
// Note: Using plain 'string' types (not union | string) for full TypeScript 3.3
// compatibility with SPFx 1.11's rush-stack-compiler-3.3.

export interface IProjectManager {
  Title: string;
  EMail: string;
}

export interface IProject {
  Id: number;
  Title: string;
  Project_Status: string;   // e.g. 'Active' | 'On Hold' | 'Completed' | 'Cancelled'
  Project_Manager: IProjectManager;
}

export interface IProjectLookup {
  Id: number;
  Title: string;
}

export interface IMilestone {
  Id: number;
  Title: string;
  DueDate: string;
  Status: string;           // e.g. 'Not Started' | 'In Progress' | 'Completed' | 'Overdue'
  ProjectLookup: IProjectLookup;
}

export interface IRisk {
  Id: number;
  Title: string;
  Impact: string;           // e.g. 'Low' | 'Medium' | 'High' | 'Critical'
  Mitigation_Plan: string;
  ProjectLookup: IProjectLookup;
}

// ─── OData Response Wrapper ───────────────────────────────────────────────────

interface IODataResponse<T> {
  value: T[];
}

// ─── Service Class ────────────────────────────────────────────────────────────

export class SharePointListService {
  private readonly context: WebPartContext;
  private readonly webAbsoluteUrl: string;

  constructor(context: WebPartContext) {
    this.context = context;
    this.webAbsoluteUrl = context.pageContext.web.absoluteUrl;
  }

  public getProjects(): Promise<IProject[]> {
    const endpoint =
      this.webAbsoluteUrl +
      "/_api/web/lists/getbytitle('PMO_Projects')/items" +
      '?$select=Id,Title,Project_Status,Project_Manager/Title,Project_Manager/EMail' +
      '&$expand=Project_Manager' +
      '&$orderby=Title asc' +
      '&$top=500';
    return this._get<IProject>(endpoint);
  }

  public getMilestones(): Promise<IMilestone[]> {
    const endpoint =
      this.webAbsoluteUrl +
      "/_api/web/lists/getbytitle('PMO_Milestones')/items" +
      '?$select=Id,Title,DueDate,Status,ProjectLookup/Id,ProjectLookup/Title' +
      '&$expand=ProjectLookup' +
      '&$orderby=DueDate asc' +
      '&$top=500';
    return this._get<IMilestone>(endpoint);
  }

  public getRisks(): Promise<IRisk[]> {
    const endpoint =
      this.webAbsoluteUrl +
      "/_api/web/lists/getbytitle('PMO_Risks')/items" +
      '?$select=Id,Title,Impact,Mitigation_Plan,ProjectLookup/Id,ProjectLookup/Title' +
      '&$expand=ProjectLookup' +
      '&$orderby=Impact desc' +
      '&$top=500';
    return this._get<IRisk>(endpoint);
  }

  private _get<T>(endpoint: string): Promise<T[]> {
    return this.context.spHttpClient
      .get(endpoint, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse): Promise<IODataResponse<T>> => {
        if (!response.ok) {
          throw new Error(
            'SharePoint REST API error [' + response.status + ']: ' +
            response.statusText
          );
        }
        return response.json();
      })
      .then((data: IODataResponse<T>): T[] => {
        return data.value;
      });
  }
}
