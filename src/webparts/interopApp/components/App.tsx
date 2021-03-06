//import { Agenda, Login, FileList, Get, MgtTemplateProps, PeoplePicker} from '@microsoft/mgt-react';
import { Login, PeoplePicker, FileList, Get, MgtTemplateProps } from '@microsoft/mgt-react/dist/es6/spfx';
import { Grid, Card, CardHeader, CardBody, Flex, Image, Text, Button, Header, Avatar, ItemLayout } from "@fluentui/react-northstar";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as React from "react";
import { useState, useEffect } from 'react';

import { Providers, ProviderState } from '@microsoft/mgt-element';


function useIsSignedIn(): [boolean] {
  const [isSignedIn, setIsSignedIn] = useState(true);
  const provider = Providers.globalProvider;
  
  useEffect(() => {
    const updateState = () => {
      const provider = Providers.globalProvider;
      setIsSignedIn(provider && provider.state === ProviderState.SignedIn);
    };

    Providers.onProviderUpdated(updateState);
    updateState();

    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    }
  }, []);

  return [isSignedIn];
}

function App() {
  const [isSignedIn] = [true];//useIsSignedIn();

  return (
    <div className="App">
      <header>
        <Login />
      </header>
      <div>
            <PeoplePicker></PeoplePicker>
        </div>
        <div>
        <ul className="breadcrumb" id="nav">
            <li><a id="home">Files</a></li>
        </ul>
        </div>
      <div>
        {isSignedIn &&
          <FileList driveId="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
          itemPath="/" enableFileUpload></FileList>
          }
      </div> 
      <div>
      {isSignedIn &&
            <Get resource="/sites?search=contoso" scopes={['Sites.Read.All']} maxPages={2}>
                    <SiteResult template="value" />
            </Get>
        }
      </div>     
    </div>
  );
}

const SiteResult = (props: MgtTemplateProps) => {
    const site = props.dataContext as MicrosoftGraph.Site;

    return (
        <div>
        <Flex gap="gap.medium" padding="padding.medium" debug>
        <Flex.Item size="size.medium">
          <div
              style={{
              position: 'relative',
              }}
          >
              <Image
              height={40}
              width={40}
              fluid
              src="https://upload.wikimedia.org/wikipedia/commons/thumb/e/e1/Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg/2097px-Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg.png"
              />
          </div>
          </Flex.Item>
          <Flex.Item grow>
          <Flex column gap="gap.small" vAlign="stretch">
              <Flex space="between">
              <Header as="h3" content={site.displayName} />
              <Text as="pre" content={site.name} />
              </Flex>

              <Text content={site.webUrl} />

              <Flex.Item push>
              <Text as="pre" content="COPYRIGHT: Fluent UI." />
              </Flex.Item>
          </Flex>
          </Flex.Item>                    
        </Flex>
      </div>
      );
    };

export default App;
