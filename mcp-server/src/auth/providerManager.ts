import { TokenManager } from "./tokenManager.js";
import { GoogleTokenManager } from "./googleTokenManager.js";
import { MicrosoftProvider } from "../providers/microsoft.js";
import { GoogleProvider } from "../providers/google.js";
import { MailProvider, AccountInfo } from "../types/mail.js";

export interface MultiAccountConfig {
  accounts: Array<{
    id: string;
    name: string;
    type: "outlook" | "gmail";
    email?: string;
    mailbox?: string; // For shared mailboxes
    clientId?: string;
    clientSecret?: string;
    tenantId?: string;
    redirectUri?: string;
  }>;
  common: {
    outlook?: {
      clientId: string;
      clientSecret?: string;
      tenantId: string;
      redirectUri: string;
    };
    gmail?: {
      clientId: string;
      clientSecret: string;
      redirectUri: string;
    };
    draftsFolderName: string;
  };
}

export class ProviderManager {
  private config: MultiAccountConfig;
  private outlookManagers: Map<string, TokenManager> = new Map();
  private gmailManagers: Map<string, GoogleTokenManager> = new Map();

  constructor(config: MultiAccountConfig) {
    this.config = config;
  }

  async getProvider(accountId?: string, overrideMailbox?: string): Promise<{ provider: MailProvider, account: AccountInfo }> {
    // If no accountId, use the first one
    const accountCfg = accountId 
      ? this.config.accounts.find(a => a.id === accountId)
      : this.config.accounts[0];

    if (!accountCfg) {
      throw new Error(`Account not found: ${accountId}`);
    }

    const mailbox = overrideMailbox || accountCfg.mailbox;

    if (accountCfg.type === "outlook") {
      const msConfig = {
        clientId: accountCfg.clientId || this.config.common.outlook?.clientId || "",
        tenantId: accountCfg.tenantId || this.config.common.outlook?.tenantId || "organizations",
        clientSecret: accountCfg.clientSecret || this.config.common.outlook?.clientSecret,
        redirectUri: accountCfg.redirectUri || this.config.common.outlook?.redirectUri || "http://localhost:3000/callback",
        scopes: ["User.Read", "Mail.Read", "Mail.ReadWrite", "Mail.Read.Shared", "Mail.ReadWrite.Shared", "Calendars.ReadWrite"],
      };

      // We use clientId + tenantId as the key for managers to reuse tokens if possible
      const managerKey = `${msConfig.clientId}-${msConfig.tenantId}`;
      let manager = this.outlookManagers.get(managerKey);
      if (!manager) {
        manager = new TokenManager(msConfig);
        this.outlookManagers.set(managerKey, manager);
      }

      const client = await manager.getClient();
      const provider = new MicrosoftProvider(client, this.config.common.draftsFolderName, mailbox);
      
      return { 
        provider, 
        account: { 
          id: accountCfg.id, 
          name: accountCfg.name, 
          email: mailbox || accountCfg.email || "", 
          type: "outlook" 
        } 
      };
    } else {
      const gConfig = {
        clientId: accountCfg.clientId || this.config.common.gmail?.clientId || "",
        clientSecret: accountCfg.clientSecret || this.config.common.gmail?.clientSecret || "",
        redirectUri: accountCfg.redirectUri || this.config.common.gmail?.redirectUri || "http://localhost:3000/callback",
        scopes: ["https://www.googleapis.com/auth/gmail.modify"],
      };

      const managerKey = gConfig.clientId;
      let manager = this.gmailManagers.get(managerKey);
      if (!manager) {
        manager = new GoogleTokenManager(gConfig);
        this.gmailManagers.set(managerKey, manager);
      }

      const gmail = await manager.getClient();
      const provider = new GoogleProvider(gmail, this.config.common.draftsFolderName);

      return { 
        provider, 
        account: { 
          id: accountCfg.id, 
          name: accountCfg.name, 
          email: accountCfg.email || "", 
          type: "gmail" 
        } 
      };
    }
  }

  getAccounts(): AccountInfo[] {
    return this.config.accounts.map(a => ({
      id: a.id,
      name: a.name,
      email: a.mailbox || a.email || "",
      type: a.type
    }));
  }
}
