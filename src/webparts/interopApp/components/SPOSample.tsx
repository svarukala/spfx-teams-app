import * as React from 'react';
import { Component } from 'react';
import * as msal from "@azure/msal-browser";
import { List } from 'office-ui-fabric-react';


interface TitleProps {
  title: string;
  subtitle?: string;
  useremail?: string;
}

let currentAccount: msal.AccountInfo = null;

const msalConfig = {  
    auth: {  
      clientId: 'c613e0d1-161d-4ea0-9db4-0f11eeabc2fd',  
      redirectUri: 'https://m365x229910.sharepoint.com/_layouts/15/workbench.aspx'  
    }  
  };
  
const msalInstance = new msal.PublicClientApplication(msalConfig); 


export default class SPOSample extends React.Component<TitleProps> {
    /*
    state = {
        resultitems: [{ name: 'Foo' }, { name: 'Bar' }]
    };
    */
constructor(props) {
    super(props);
    this.state = {  
        resultitems: [  
          {  
            "RefinableString10": "",  
            "CreatedBy": "",  
            "Created":""  
         
          }  
        ]  
      };  
    }
public componentDidMount(){

    this.setCurrentAccount();
    msalInstance.acquireTokenSilent(this.tokenrequest).then((val) => {  
    let headers = new Headers();  
    let bearer = "Bearer " + val.accessToken;  
    console.info("BEARER TOKEN: "+ val.accessToken);
    headers.append("Authorization", bearer); 
    headers.append("Accept", "application/json;odata=verbose");
    headers.append("Content-Type", "application/json;odata=verbose");
    let options = {  
        method: "GET",  
        headers: headers  
    };  
    fetch("https://m365x229910.sharepoint.com/sites/DevDemo/_api/search/query?querytext='*'", options)
        .then(resp => {  
            resp.json().then((data) => {  
                console.log(data);
                //var jsonObject = JSON.parse(data.body);
                this.setState({
                    resultitems: data.PrimaryQueryResult.RelevantResults.Table.Rows
                }); 
            });
        });  
    }).catch((errorinternal) => {  
        console.log(errorinternal);  
    });  

}
public tokenrequest: msal.SilentRequest = {
    scopes: ["https://m365x229910.sharepoint.com/AllSites.Read", "https://m365x229910.sharepoint.com/AllSites.Manage"],
    account: currentAccount,
    };     

private getSites = ():void => {
    this.setCurrentAccount();
    msalInstance.acquireTokenSilent(this.tokenrequest).then((val) => {  
    let headers = new Headers();  
    let bearer = "Bearer " + val.accessToken;  
    console.info("BEARER TOKEN: "+ val.accessToken);
    headers.append("Authorization", bearer); 
    headers.append("Accept", "application/json;odata=verbose");
    headers.append("Content-Type", "application/json;odata=verbose");
    let options = {  
        method: "GET",  
        headers: headers  
    };  
    fetch("https://m365x229910.sharepoint.com/sites/DevDemo/_api/search/query?querytext='*'", options)
        .then(resp => {  
            resp.json().then((data) => {  
                console.log(data);
                //var jsonObject = JSON.parse(data.body);
                var tempItems = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results
                

                this.setState({resultitems:tempItems});
                //return results;  
            });
        });  
    }).catch((errorinternal) => {  
        console.log(errorinternal);  
    });  
}
protected setCurrentAccount = (): void => {
    const currentAccounts: msal.AccountInfo[] = msalInstance.getAllAccounts();
    if (currentAccounts === null || currentAccounts.length == 0) {
        this.tokenrequest.account = msalInstance.getAccountByUsername(
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
    this.tokenrequest.account = currentAccount;
  }; 

  render() {
    const { title, subtitle, children, useremail } = this.props;
    //this.getSites();
    return (
      <>
        <h1>{title}</h1>
        <h2>{subtitle}</h2>
        <h2>{useremail}</h2>
        <div>{children}</div>
        
      </>
    );
  }
}

//export default SPOSample;