/* eslint-disable @typescript-eslint/no-explicit-any */

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export default class GraphService {
  private context: BaseComponentContext;
  private client: MSGraphClientV3 | undefined;

  constructor(context: BaseComponentContext) {
    this.context = context;
  }

  private async getClient(): Promise<MSGraphClientV3> {
    if (!this.client) {
      this.client = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.client;
  }


 public async getUserGroups(): Promise<any[]> {
    const client = await this.getClient();
    const result = await client.api('/me/memberOf').get();

    console.log("GraphService - getUserGroups result:", result);

    return result.value;
  }



}