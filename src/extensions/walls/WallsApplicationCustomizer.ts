import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import { graph } from "@pnp/graph/presets/all";
import "@pnp/graph/users";

import * as strings from 'WallsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'WallsApplicationCustomizer';

export interface IWallsApplicationCustomizerProperties {
}

export default class WallsApplicationCustomizer
  extends BaseApplicationCustomizer<IWallsApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    var walls = await this._checkUser();

    if(!walls) {
      this.context.application.navigatedEvent.add(this, this._render);
    }

    return Promise.resolve();
  }

  public async _checkUser() {
    graph.setup({
      spfxContext: this.context
    });

    let isAdmin = false;

    let user: any[] = await graph.me.memberOf();

    for(let groups of user) {
      if(groups.roleTemplateId && groups.roleTemplateId === "f28a1f50-f6e7-4571-818b-6a12f2af6b6c") { // Sharepoint
        isAdmin = true;
      } else if(groups.id === "315f2b29-7a6d-4715-b3cf-3af28d0ddf4b") { // UX DESIGN
        isAdmin = true;
      } else if(groups.id === "24998f56-6911-4041-b4d1-f78452341da6") { // Support
        isAdmin = true;
      }
    }

    return isAdmin;
  }

  public _render(){
    // set interval
    this._setSettingsPaneInterval();

    // Site contents page
    if (this.context.pageContext.site.serverRequestPath === "/_layouts/15/viewlsts.aspx") {
      window.setTimeout(() => {
        let commandBar = document.querySelector(".ms-CommandBar-secondaryCommand");

        let wF = commandBar.querySelectorAll('button[name="Site workflows"]');
        wF[0].remove();
        let sS = commandBar.querySelectorAll('button[name="Site settings"]');
        sS[0].remove();
      }, 175);
    }
  }

  // Check for settings pane
  public async _setSettingsPaneInterval(){
    let interval = setInterval(() => {
      var settingsPane = document.getElementById('FlexPane_Settings');

      if(settingsPane) {
        this._addWalls(settingsPane);

        // No more searching
        clearInterval(interval);
        this._setSettingsRemoveInterval();

      }


    }, 500);
  }

  // See if settings pane has been closed
  public async _setSettingsRemoveInterval(){
    let interval = setInterval(() => {
      var settingsPane = document.getElementById('FlexPane_Settings');

      if(!settingsPane) {

        // No more searching
        clearInterval(interval);
        this._setSettingsPaneInterval();
      }


    }, 600);
  }

  public async _addWalls(settingsPane) {
    // Remove options in settings
    // Site permissions
    var sP = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSitePermissions();"]');
    if(sP.length > 0) sP[0].remove();
    sP = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SitePermissions");
    if(sP.length > 0) sP[0].remove();

    // Site information
    var sI = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteSettings();"]');
    if(sI.length > 0) {
      let element: HTMLElement = sI[0] as HTMLElement;
      element.onclick = () => {
        window.setTimeout(() => {
          var siteSettingsPane = document.getElementsByClassName("ms-SiteSettingsPanel-SiteInfo");
          if(siteSettingsPane.length > 0) {
            window.setTimeout(() => {
            var jhs = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-joinHubSite");
            if(jhs.length >0 ) jhs[0].remove();
            }, 300);
            var c = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-classification");
            if(c.length >0 ) c[0].remove();
            var p = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-PrivacyDropdown");
            if(p.length >0 ) p[0].remove();
            var ht = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-HelpText");
            if(ht.length >0 ) ht[0].remove();
          }
        }, 500);
      }
    }

    var sI2 = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SiteInformation");
    if(sI2.length > 0) sI2[0].remove();

    // Change the look
    var cTL = settingsPane.querySelectorAll('a[href="javascript:_spLaunchChangeTheLookPanel();"]');
    if(cTL.length > 0) cTL[0].remove();
    cTL = settingsPane.querySelectorAll("#Change_The_Look");
    if(cTL.length > 0) cTL[0].remove();

    // Site Designs
    var sD = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteDesignProgress();"]');
    if(sD.length > 0) sD[0].remove();
    sD = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SiteDesigns");
    if(sD.length > 0) sD[0].remove();
  }
}
