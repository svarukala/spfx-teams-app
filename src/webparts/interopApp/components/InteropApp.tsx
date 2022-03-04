import * as React from 'react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import styles from './InteropApp.module.scss';
import { IInteropAppProps } from './IInteropAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Person, PeoplePicker, FileList, Get, MgtTemplateProps } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType} from '@microsoft/mgt-spfx';
import { Pivot, PivotItem, Image, List, ImageFit } from 'office-ui-fabric-react';
//import SPOSample from './SPOSample';
import SPOSearch from './SPOSearch';
import ReusableApp from './ReusableApp';

export default class InteropApp extends React.Component<IInteropAppProps, {}> {
  public render(): React.ReactElement<IInteropAppProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      userEmail,
      idToken
    } = this.props;

    console.log("idToken: " + idToken);
    return (
      <section className={`${styles.interopApp} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Welcome, {escape(userDisplayName)}!</h2>
        </div>
        <div>
          <div>
            <Person personQuery="me" view={ViewType.image}></Person>
          </div>

          <Pivot aria-label="Basic Pivot Example">
          <PivotItem headerText="Files">
            <FileList></FileList> 
          </PivotItem>
          <PivotItem headerText="People">
            <br/>
            <PeoplePicker></PeoplePicker>
            </PivotItem>
            <PivotItem headerText="File Upload">
            <FileList driveId="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
          itemPath="/" enableFileUpload></FileList>
              </PivotItem>
          
          <PivotItem headerText="Sites Search Using MSGraph">
              <Get resource="/sites?search=contoso" scopes={['Sites.Read.All']} maxPages={2}>
                      <SiteResult template="value" />
              </Get>
          </PivotItem>

          <PivotItem headerText="Sites Search Using SPO REST API">
            {/*<SPOSample title="Hello world" subtitle="Welcome!" useremail={userEmail} ></SPOSample>*/}
            <SPOSearch useremail={userEmail}></SPOSearch>
          </PivotItem>
          <PivotItem headerText="SPO Search">
              <ReusableApp idToken={idToken} />
            </PivotItem>
          </Pivot>
        </div>
      </section>
    );
  }
}


const SiteResult = (props: MgtTemplateProps) => {
  const site = props.dataContext as MicrosoftGraph.Site;

  return (

    <div className="ms-ListBasicExample-itemCell">
    <Image
      className="ms-ListBasicExample-itemImage"
      src="https://upload.wikimedia.org/wikipedia/commons/thumb/e/e1/Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg/2097px-Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg.png"
      width={50}
      height={50}
      imageFit={ImageFit.cover}
    />
      <div className='site'>
          <div className="title">
					<a href={site.webUrl??""} target="_blank" rel="noreferrer">
						<h3>{site.displayName}</h3>
					</a>
					<span className="date">
						{new Date(site.createdDateTime??"").toLocaleDateString()}
					</span>
				</div>
    </div>
    </div>

    );
  };