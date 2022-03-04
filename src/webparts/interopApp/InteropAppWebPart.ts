import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import App from './components/App';
import * as strings from 'InteropAppWebPartStrings';
import InteropApp from './components/InteropApp';
import { IInteropAppProps } from './components/IInteropAppProps';
import { Providers, SharePointProvider, SimpleProvider, ProviderState, Graph } from '@microsoft/mgt-spfx';
import { IPublicClientApplication, Configuration, LogLevel, PublicClientApplication, AccountInfo, SilentRequest, 
  InteractionRequiredAuthError, AuthorizationUrlRequest, InteractionRequiredAuthErrorMessage } from "@azure/msal-browser";

export interface IInteropAppWebPartProps {
  description: string;
}

const msalConfig: Configuration = {
  auth: {
    clientId: "c613e0d1-161d-4ea0-9db4-0f11eeabc2fd",
    authority: "https://login.microsoftonline.com/044f7a81-1422-4b3d-8f68-3001456e6406",
    redirectUri:"https://m365x229910.sharepoint.com/sites/DevDemo/_layouts/15/workbench.aspx",
  },
  system: {
    iframeHashTimeout: 10000,
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
    },
  },
};

const msalInstance: PublicClientApplication = new PublicClientApplication(
  msalConfig
);

let currentAccount: AccountInfo = null;
let idTokenVal: string = null;

const tokenrequest: SilentRequest = {
  scopes: ['Mail.Read','calendars.read', 'user.read', 'openid', 'profile', 'people.read', 'user.readbasic.all', 'files.read', 'files.read.all'],
  account: currentAccount,
}; 

export default class InteropAppWebPart extends BaseClientSideWebPart<IInteropAppWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    if (!Providers.globalProvider) {
      console.log('Initializing global provider');
      Providers.globalProvider = new SimpleProvider(async ()=>{console.log("For simple provider"); return this.getAccessToken();}); //, async ()=>{ Providers.globalProvider.setState(ProviderState.SignedIn)}, async ()=>{});
      //new SharePointProvider(this.context);
      Providers.globalProvider.setState(ProviderState.SignedIn);
      this.getAccessToken();
    }
    
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IInteropAppProps> = React.createElement(
      InteropApp,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        idToken: "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImtpZCI6Ik1yNS1BVWliZkJpaTdOZDFqQmViYXhib1hXMCJ9.eyJhdWQiOiJjNjEzZTBkMS0xNjFkLTRlYTAtOWRiNC0wZjExZWVhYmMyZmQiLCJpc3MiOiJodHRwczovL2xvZ2luLm1pY3Jvc29mdG9ubGluZS5jb20vMDQ0ZjdhODEtMTQyMi00YjNkLThmNjgtMzAwMTQ1NmU2NDA2L3YyLjAiLCJpYXQiOjE2NDYzNjY1NTEsIm5iZiI6MTY0NjM2NjU1MSwiZXhwIjoxNjQ2MzcwNDUxLCJhaW8iOiJBVFFBeS84VEFBQUF4VEh5V0hUKzBld25sbDk5R2J3b2JPSWU4UnRYSkNnK1JCYStnUFoxNVA4SUwwVWEzSzFhNTlQaUZpL3RmZzhPIiwibmFtZSI6Ik1lZ2FuIEJvd2VuIiwibm9uY2UiOiI1ZTY4NGEyZC0wOTMxLTRjZWMtYjIxYy1mNmFjZGMwNjYwMjYiLCJvaWQiOiI0Y2IwOGRjYi1iNTBlLTRlZTYtOTcxMi0wM2ZkNGM3NDZhNmMiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiJNZWdhbkJATTM2NXgyMjk5MTAuT25NaWNyb3NvZnQuY29tIiwicmgiOiIwLkFUY0FnWHBQQkNJVVBVdVBhREFCUlc1a0J0SGdFOFlkRnFCT25iUVBFZTZyd3YwM0FITS4iLCJzdWIiOiJSTzNrWmliRHFjZ3oyZUJPQUt4VUd6R2RIUkplU3pORFFaTGUyT1h5WVV3IiwidGlkIjoiMDQ0ZjdhODEtMTQyMi00YjNkLThmNjgtMzAwMTQ1NmU2NDA2IiwidXRpIjoiSmY2ZEVYakJVMEtKcjhUcnUyaWtBQSIsInZlciI6IjIuMCJ9.MaRvESj2bCFrc7STnOrPtC-4S-4fgA_BIbr5qsV6SG_rOa8kHLUcK4takQMz8-CjNE41dcoQa5t2u_cLFf0dUxT3oio-kSP3aG55O-27V5TIF-t5i8ogX9qQPJ3cD0guZeWrHjIXkqh87E9lPhxVWUT0CfkKlwOtz30VpoOVNZYQ_nwNqh9zGUbdcVXz0ASZFLwcr2GRt6Dz8WdCX2DgIyDPSU6oUTPDAHyNAAqP8Oa2YYUMzGsEFPXZN63Jiqa66LEV_piq4ZvYp6lRSVkYq025LCcDQpSlLVwlai5vC-ZxH0vwokp7qOGSDQ77oC4xiVSuU7oOPW0tGj3mSsAqHg"
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /*
  public render(): void {
    this.domElement.innerHTML = `
            <mgt-login></mgt-login>
            <div>
            <mgt-file-list drive-id="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
          item-path="/" enable-file-upload></mgt-file-list>
            </div>
            <div>
              <p>Email Subject:<span id="email"></span> </p>
              <p>Error, if any : <span id="error"></span></p>
            </div>
            `;
    this._searchWithGraph();
  } 

  public render(): void {
    this.domElement.innerHTML = `
      <mgt-agenda></mgt-agenda>
      `;
  }
  
  public render(): void {
    const element: React.ReactElement<IInteropAppProps> = React.createElement(
      InteropApp,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  const exchangeSsoTokenForOboTokenForSPO = async () => {
  const response = await fetch(`http://localhost:5000/api/token?ssoToken=abc&tokenFor=spo`, { method: 'POST' });
  const responsePayload = await response.json();
  if (response.ok) {
    //setSPOOboToken(responsePayload.access_token);
    console.log(responsePayload);
  } else {
    if (responsePayload!.error === "consent_required") {
      //setError("consent_required");
      console.log("consent_required");
    } else {
      //setError("unknown SSO error");
      console.log("unknown SSO error");
    }
  }
};
  */
  
  protected setCurrentAccount = (): void => {
    const currentAccounts: AccountInfo[] = msalInstance.getAllAccounts();
    if (currentAccounts === null || currentAccounts.length == 0) {
      currentAccount = msalInstance.getAccountByUsername(
        this.context.pageContext.user.loginName
      );
    } else if (currentAccounts.length > 1) {
      console.warn("Multiple accounts detected.");
      currentAccount = msalInstance.getAccountByUsername(
        this.context.pageContext.user.loginName
      );
    } else if (currentAccounts.length === 1) {
      currentAccount = currentAccounts[0];
    }
    tokenrequest.account = currentAccount;
  }; 

  protected getAccessToken = async (): Promise<string> => {
    console.log("Getting access token");
    let accessToken: string = null;
    this.setCurrentAccount();
    console.log(currentAccount);
    return msalInstance
      .acquireTokenSilent(tokenrequest)
      .then((tokenResponse) => {
        console.log("Inside Silent");
        console.log("Access token: "+ tokenResponse.accessToken);
        console.log("ID token: "+ tokenResponse.idToken);
        idTokenVal = tokenResponse.idToken;
        return tokenResponse.accessToken;
      })
      .catch((err) => {
        console.log(err);
        console.log("Silent Failed");
        if (err instanceof InteractionRequiredAuthError) {
          return this.interactionRequired();
        } else {
          console.log("Some other error. Inside SSO.");
          //const loginPopupRequest: AuthorizationUrlRequest = tokenrequest;
          const loginPopupRequest: AuthorizationUrlRequest = tokenrequest as AuthorizationUrlRequest;
          loginPopupRequest.loginHint = this.context.pageContext.user.loginName;
          return msalInstance
            .ssoSilent(loginPopupRequest)
            .then((tokenResponse) => {
              idTokenVal = tokenResponse.idToken;
              return tokenResponse.accessToken;
            })
            .catch((ssoerror) => {
              console.error(ssoerror);
              console.error("SSO Failed");
              if (ssoerror) {
                return this.interactionRequired();
              }
              return null;
            });
        }
      });
  };

  protected interactionRequired = (): Promise<string> => {
    console.log("Inside Interaction");
    const loginPopupRequest: AuthorizationUrlRequest = tokenrequest as AuthorizationUrlRequest;
    loginPopupRequest.loginHint = this.context.pageContext.user.loginName;
    return msalInstance
      .acquireTokenPopup(loginPopupRequest)
      .then((tokenResponse) => {
        return tokenResponse.accessToken;
      })
      .catch((error) => {
        console.error(error);
        // I haven't implemented redirect but it is fairly easy
        console.error("Maybe it is a popup blocked error. Implement Redirect");
        return null;
      });
  }; 

  protected _searchWithGraph = (): void => {
    // Log the current operation
    console.log("Using _searchWithGraph() method");
    this.getAccessToken().then((accessToken) => {
      if (accessToken != null) {
        const headers = new Headers();
        const bearer = `Bearer ${accessToken}`;
        headers.append("Authorization", bearer);
        const options = {
          method: "GET",
          headers: headers,
        };
        console.log("request made to Graph API at: " + new Date().toString());
        fetch("https://graph.microsoft.com/v1.0/me/messages", options)
          .then((response) => response.json())
          .then((data) => {
            console.log(data);
            document.getElementById("email").innerText = data.value
              .map((o) => o.subject)
              .join(" || ");
          })
          .catch((error) => console.log(error));
      } else {
        document.getElementById("error").innerText =
          "Error! Check browser console";
      }
    });
  }; 

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

