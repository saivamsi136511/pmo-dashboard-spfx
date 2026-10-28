import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// ─── Data Model Interfaces ────────────────────────────────────────────────────

export interface IProjectManager {
  Title: string;
  EMail?: string;
}

export interface IProject {
  Id: number;
  Title: string;
  Project_Status: 'Active' | 'On Hold' | 'Completed' | 'Cancelled' | string;
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
  Status: 'Not Started' | 'In Progress' | 'Completed' | 'Overdue' | string;
  ProjectLookup: IProjectLookup;
}

export interface IRisk {
  Id: number;
  Title: string;
  Impact: 'Low' | 'Medium' | 'High' | 'Critical' | string;
  Mitigation_Plan: string;
  ProjectLookup: IProjectLookup;
}

// ─── OData Response Wrappers ──────────────────────────────────────────────────

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

  /**
   * Fetch all projects with the Project_Manager person field expanded.
   */
  public async getProjects(): Promise<IProject[]> {
    const endpoint =
      `${this.webAbsoluteUrl}/_api/web/lists/getbytitle('PMO_Projects')/items` +
      `?$select=Id,Title,Project_Status,Project_Manager/Title,Project_Manager/EMail` +
      `&$expand=Project_Manager` +
      `&$orderby=Title asc` +
      `&$top=500`;

    return this._get<IProject>(endpoint);
  }

  /**
   * Fetch all milestones with the ProjectLookup expanded to surface parent project info.
   */
  public async getMilestones(): Promise<IMilestone[]> {
    const endpoint =
      `${this.webAbsoluteUrl}/_api/web/lists/getbytitle('PMO_Milestones')/items` +
      `?$select=Id,Title,DueDate,Status,ProjectLookup/Id,ProjectLookup/Title` +
      `&$expand=ProjectLookup` +
      `&$orderby=DueDate asc` +
      `&$top=500`;

    return this._get<IMilestone>(endpoint);
  }

  /**
   * Fetch all risks with the ProjectLookup expanded to surface parent project info.
   */
  public async getRisks(): Promise<IRisk[]> {
    const endpoint =
      `${this.webAbsoluteUrl}/_api/web/lists/getbytitle('PMO_Risks')/items` +
      `?$select=Id,Title,Impact,Mitigation_Plan,ProjectLookup/Id,ProjectLookup/Title` +
      `&$expand=ProjectLookup` +
      `&$orderby=Impact desc` +
      `&$top=500`;

    return this._get<IRisk>(endpoint);
  }

  // ─── Private Helpers ────────────────────────────────────────────────────────

  private async _get<T>(endpoint: string): Promise<T[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(
        `SharePoint REST API error [${response.status}]: ${response.statusText} — ${endpoint}`
      );
    }

    const data: IODataResponse<T> = await response.json();
    return data.value;
  }
}
