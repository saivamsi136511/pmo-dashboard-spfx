import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class SharePointListService {
  private context: WebPartContext;
  private absoluteUrl: string;

  constructor(context: WebPartContext) {
    this.context = context;
    this.absoluteUrl = context.pageContext.web.absoluteUrl;
  }

  // Fetch Projects
  public async getProjects(): Promise<any[]> {
    const endpoint = `${this.absoluteUrl}/_api/web/lists/getbytitle('PMO_Projects')/items?$select=Id,Title,Project_Status,Project_Manager/Title,Description&$expand=Project_Manager`;
    return this.get(endpoint);
  }

  // Fetch Milestones with Lookup Expansion to PMO_Projects
  public async getMilestones(): Promise<any[]> {
    const endpoint = `${this.absoluteUrl}/_api/web/lists/getbytitle('PMO_Milestones')/items?$select=Id,Title,DueDate,Status,ProjectLookup/Title,ProjectLookup/Id&$expand=ProjectLookup`;
    return this.get(endpoint);
  }

  // Fetch Risks with Lookup Expansion to PMO_Projects
  public async getRisks(): Promise<any[]> {
    const endpoint = `${this.absoluteUrl}/_api/web/lists/getbytitle('PMO_Risks')/items?$select=Id,Title,Impact,Mitigation_Plan,ProjectLookup/Title,ProjectLookup/Id&$expand=ProjectLookup`;
    return this.get(endpoint);
  }

  // Generic GET
  private async get(endpoint: string): Promise<any[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      throw new Error(`Error fetching data: ${response.statusText}`);
    }
    const json = await response.json();
    return json.value;
  }

  // Generic POST for creating items
  public async createItem(listName: string, payload: any): Promise<any> {
    const endpoint = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    const options: ISPHttpClientOptions = {
      body: JSON.stringify(payload)
    };

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Error creating item: ${response.statusText}`);
    }
    return await response.json();
  }
}