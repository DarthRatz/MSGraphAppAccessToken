import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import axios from "axios";
import { stringify } from "qs";
import { MicrosoftAppDetails } from "./MicrosoftAppDetails";
require("isomorphic-fetch");

interface OathToken {
  client_id: string;
  client_secret: string;
  scope: string;
  grant_type: string;
}

export class ClientCredentialAuthenticationProvider
  implements AuthenticationProvider {
  public async getAccessToken(): Promise<string> {
    const url: string = `https://login.microsoftonline.com/${MicrosoftAppDetails.TenantId}/oauth2/v2.0/token`;

    const body: OathToken = {
      client_id: MicrosoftAppDetails.AppId,
      client_secret: MicrosoftAppDetails.Password,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    };

    try {
      let response = await axios.post(url, stringify(body));

      if (response.status === 200) {
        return response.data.access_token;
      } else {
        throw new Error("Non 200OK response on obtaining token...");
      }
    } catch (error) {
      throw new Error("Error on obtaining token...");
    }
  }
}
