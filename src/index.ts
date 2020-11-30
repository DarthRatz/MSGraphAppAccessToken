import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { ClientCredentialAuthenticationProvider } from "./AuthenticationProvider";

// Create Graph SDK Client
function createAuthenticatedClient(): Client {
  const clientOptions: ClientOptions = {
    defaultVersion: "v1.0",
    debugLogging: false,
    authProvider: new ClientCredentialAuthenticationProvider(),
  };

  const client: Client = Client.initWithMiddleware(clientOptions);

  return client;
}

// Get Users from Graph
async function getUsers(): Promise<any> {
  const client: Client = createAuthenticatedClient();

  const request = await client
    .api("/users")
    .select("id, displayName, mail")
    .get()
    .catch((error) => {
      console.log(error);
    });

  console.log(request.value);
}

getUsers();
