// Gets graph access token using `client_credentials` flow
// TODO: Implement cache to reuse the short lived access token(within ~1hr)

import axios, { AxiosRequestConfig, AxiosPromise } from "axios";
import { stringify } from 'qs';

const AUTH_BASE_URL = "https://login.microsoftonline.com";

export default class Auth {
  clientId: string;
  clientSecret: string;
  tenantId: string;

  constructor(clientId: string, clientSecret: string, tenantId: string) {
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.tenantId = tenantId;
  }

  /**
   * @param refresh: boolean to tell if the token needs to be refreshed from onedrive
   */
  async getAccessToken(refresh: boolean): Promise<string> {
    let token: string;
    if (refresh) {
      const data = {
        client_id: this.clientId,
        client_secret: this.clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials"
      }
      const response = await axios.post(`${AUTH_BASE_URL}/${this.tenantId}/oauth2/v2.0/token`, stringify(data))
      token = response.data.access_token;
      // saveToCache(token);
    } else {
      token = await this.readFromCache();
      if (!token) {
        return this.getAccessToken(true);
      }
    }
    return token;
  }

  async saveToCache(token: string) {

  }

  async readFromCache(): Promise<string> {
    return "todo";
  }

}