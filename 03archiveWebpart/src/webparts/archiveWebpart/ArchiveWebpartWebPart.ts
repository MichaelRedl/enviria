import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import { PermissionKind } from '@pnp/sp/security';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './ArchiveWebpartWebPart.module.scss';

export interface IArchiveWebpartWebPartProps {
  projectArchived: boolean;
}

export interface IArchiveWebpartWebPartState {
  showConfirmation: boolean;
}

export default class ArchiveWebpartWebPart extends BaseClientSideWebPart<IArchiveWebpartWebPartProps> {
  private userHasEditPermission: boolean = false;
  private sp: ReturnType<typeof spfi>;
  private projectStatus: string = '';
  private externalSiteUrl: string = '';
  private reactivateApiUrl: string = '';
  private archiveApiUrl: string = '';

  private state: IArchiveWebpartWebPartState = {
    showConfirmation: false,
  };

  protected async onInit(): Promise<void> {
    await super.onInit();
    this.properties.projectArchived = this.properties.projectArchived || false;

    // Initialize PnP JS
    this.sp = spfi().using(SPFx(this.context));

    // Determine the external site URL
    const urlName = this.context.pageContext.web.absoluteUrl.split('/').pop();
    if (urlName && urlName.toLowerCase().indexOf("project") !== -1) {
      this.externalSiteUrl = 'https://enviria.sharepoint.com/sites/Projects';
      this.reactivateApiUrl = '';
      this.archiveApiUrl = '';
    } else {
      this.externalSiteUrl = 'https://enviria.sharepoint.com/sites/Test-Projects';
      this.reactivateApiUrl = 'https://prod2-23.germanywestcentral.logic.azure.com:443/workflows/b13b2a09a71542738aefa179b4004865/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=vXUeG-DQzH6pTBIsvIQpjEIwApdC5siYzmtACu3BKZg';
      this.archiveApiUrl = 'https://prod2-08.germanywestcentral.logic.azure.com:443/workflows/d4b0382ee6444847904b87b7b45de3a7/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ljEk5tz3PP5HKYCIMy4tY880UiOsR2bUzKMUXIYED_E';
    }

    // Check if user has edit permissions
    try {
      this.userHasEditPermission = await this._checkUserPermissions();
      console.log('USER PERMISSIONS');
      console.log(this.userHasEditPermission);
    } catch (error) {
      console.error('Error checking permissions:', error);
      this.userHasEditPermission = false; // Default to false if there's an error
    }

    // Fetch project status
    try {
      this.projectStatus = await this._fetchProjectStatus();
    } catch (error) {
      console.error('Error fetching project status:', error);
      this.projectStatus = 'Unknown'; // Default to 'Unknown' if there's an error
    }

    // Set projectArchived based on projectStatus
    this.properties.projectArchived = this.projectStatus === 'Archived';

    this.render(); // Render after initialization and data fetching
  }

  private async _checkUserPermissions(): Promise<boolean> {
    try {
      const currentUser = await this.sp.web.currentUser();
      const userEffectivePermissions = await this.sp.web.getUserEffectivePermissions(currentUser.LoginName);
      let temp = this.sp.web.hasPermissions(userEffectivePermissions, PermissionKind.EditListItems);
      console.log(temp);
      return true;
      // return this.sp.web.hasPermissions(userEffectivePermissions, PermissionKind.EditListItems);
    } catch (error) {
      if (error.status === 403) {
        return false; // User does not have sufficient permissions
      } else {
        throw error; // Re-throw other errors
      }
    }
  }

  private async _fetchProjectStatus(): Promise<string> {
    const urlName = this.context.pageContext.web.absoluteUrl.split('/').pop();
    const listUrl = `${this.externalSiteUrl}/_api/web/lists/getByTitle('Project Sites')/items?$filter=URL_x0020_Name eq '${urlName}'&$select=Status`;

    try {
      const response = await this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const data = await response.json();
        if (data.value && data.value.length > 0 && data.value[0].Status) {
          return data.value[0].Status; // Assuming 'Status' is a managed metadata column
        }
      }
      return 'Unknown';
    } catch (error) {
      console.error('Error fetching project status:', error);
      return 'Unknown';
    }
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.archiveWebpart}">
        ${this.state.showConfirmation ? this._renderConfirmationPopup() : ''}
        ${this.properties.projectArchived ?
        `<div class="${styles.archiveContainer}">
            <div class="${styles.archiveMessage}">This project has been archived. Please contact the IT department if you want to reactivate it.</div>
            ${this.userHasEditPermission && this.projectStatus === 'Archived' ?
          `<button id="reactivateButton" class="${styles.reactivateButton}">Reactivate Project</button>`
          : ''}
          </div>` :
        `<div class="${styles.archiveContainer}">
            ${this.projectStatus === 'Active' ?
          `<button id="archiveButton" class="${styles.statusMessage}">Project Status: <div style="color:green">Active</div></button>`
          : ''}
          </div>`
      }
      </div>
    `;

    if (!this.properties.projectArchived && this.projectStatus === 'Active') {
      this._bindArchiveButtonEvent();
    } else if (this.properties.projectArchived && this.userHasEditPermission && this.projectStatus === 'Archived') {
      this._bindReactivateButtonEvent();
    }
  }

  private _renderConfirmationPopup(): string {
    return `
      <div class="${styles.overlay}"></div>
      <div class="${styles.confirmationPopup}">
        <p>Möchten Sie dieses Projekt wirklich archivieren?</p>
        <div style ="font-size:12px; margin-bottom: 30px">Nicht mehr bearbeitbar, weiterhin auffindbar. Reaktivierung über IT möglich.</div>
        <button id="confirmArchiveButton" class="${styles.confirmButton}">Ja</button>
        <button id="cancelArchiveButton" class="${styles.cancelButton}">Nein</button>
      </div>
    `;
  }

  private _bindArchiveButtonEvent(): void {
    const archiveButton = this.domElement.querySelector('#archiveButton');
    if (archiveButton) {
      archiveButton.addEventListener('click', () => this._onArchiveButtonClick());
    }

    const confirmButton = this.domElement.querySelector('#confirmArchiveButton');
    if (confirmButton) {
      confirmButton.addEventListener('click', () => this._onConfirmArchive());
    }

    const cancelButton = this.domElement.querySelector('#cancelArchiveButton');
    if (cancelButton) {
      cancelButton.addEventListener('click', () => this._onCancelArchive());
    }
  }

  private _bindReactivateButtonEvent(): void {
    const reactivateButton = this.domElement.querySelector('#reactivateButton');
    if (reactivateButton) {
      reactivateButton.addEventListener('click', () => this._onReactivateButtonClick());
    }
  }

  private async _onArchiveButtonClick(): Promise<void> {
    this.setState({ showConfirmation: true });
  }

  private async _onConfirmArchive(): Promise<void> {
    this.setState({ showConfirmation: false });
    this.properties.projectArchived = true;
    await this._makeArchiveApiCall();
    this.render(); // Re-render the web part to show the archived message and reactivate button
  }

  private _onCancelArchive(): void {
    this.setState({ showConfirmation: false });
  }

  private async _onReactivateButtonClick(): Promise<void> {
    this.properties.projectArchived = false;
    this.projectStatus = "Active"
    await this._makeReactivateApiCall();
    this.render(); // Re-render the web part to show the archive button
  }

  private async _makeArchiveApiCall(): Promise<void> {
    const url = this.context.pageContext.web.absoluteUrl;
    const urlName = url.split('/').pop();
    const pipedriveID = urlName ? urlName.split('-').pop() : '';

    const apiUrl = this.archiveApiUrl;

    const body = JSON.stringify({ "pipedriveID": pipedriveID });

    try {
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: body
      });

      if (!response.ok) {
        throw new Error(`API call failed with status code ${response.status}`);
      }
    } catch (error) {
      console.error('Error making API call:', error);
    }
  }

  private async _makeReactivateApiCall(): Promise<void> {
    const url = this.context.pageContext.web.absoluteUrl;
    const urlName = url.split('/').pop();
    const pipedriveID = urlName ? urlName.split('-').pop() : '';

    const apiUrl = this.reactivateApiUrl;

    const body = JSON.stringify({ "pipedriveID": pipedriveID });

    try {
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: body
      });

      if (!response.ok) {
        throw new Error(`API call failed with status code ${response.status}`);
      }
    } catch (error) {
      console.error('Error making API call:', error);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private setState(newState: Partial<IArchiveWebpartWebPartState>): void {
    this.state = { ...this.state, ...newState };
    this.render();
  }
}
