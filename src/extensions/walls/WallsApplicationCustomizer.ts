import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import { graph } from "@pnp/graph/presets/all";
import "@pnp/graph/users";
import { PermissionKind } from '@pnp/pnpjs';
import { sp } from "@pnp/sp/presets/all";

const LOG_SOURCE: string = 'WallsApplicationCustomizer';

export interface IWallsApplicationCustomizerProperties {
}

export default class WallsApplicationCustomizer
  extends BaseApplicationCustomizer<IWallsApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    var walls = await this._checkUser();
    console.log("User is "+walls);
    if (walls != "admin") {
      console.log("User is not admin");
      this.context.application.navigatedEvent.add(this, this._render);
    }

    return Promise.resolve();
  }

  public async _checkUser() {
    graph.setup({
      spfxContext: this.context
    });

    let permissions = await sp.web.getCurrentUserEffectivePermissions();
    let isOwner = false;
    let userType = "user"
    let templateType = this.context.pageContext.web.templateName; // 64: teams, 68: comms

    let user: any[] = await graph.me.memberOf();
    if (sp.web.hasPermissions(permissions, PermissionKind.ManageWeb) && sp.web.hasPermissions(permissions, PermissionKind.ManagePermissions) && sp.web.hasPermissions(permissions, PermissionKind.CreateGroups)) {
      isOwner = true// check if user is a owner by checking the permission
    }

    for (let groups of user) {
      console.log(groups.id)
      if (templateType == "64") { // If site is a teams site (no group member on comms site)
        if (groups.id === this.context.pageContext.site.group.id["_guid"]) { // If user is member of the group
          console.log("Member");
          userType = "member";
        }
      }

      if (groups.id === "c32ff810-25ae-43d3-af87-0b2b5c41dc09") { // SCA
        console.log("admin");
        userType = "admin";
      } else if (groups.id === "315f2b29-7a6d-4715-b3cf-3af28d0ddf4b") { // UX DESIGN
        userType = "admin";
        console.log("admin");
      } else if (groups.id === "24998f56-6911-4041-b4d1-f78452341da6") { // Support
        userType = "admin";
        console.log("admin");
      }
    }

    //If user is an admin, it should keep the admin access not owner
    if (isOwner && userType != "admin") { 
      userType = "owner"
      console.log("owner")
    }
    return userType;
  }

  public _render(){
    // set interval
    this._setSettingsPaneInterval();

    // Site contents page
    if (this.context.pageContext.site.serverRequestPath === "/_layouts/15/viewlsts.aspx") {
      window.setTimeout(() => {
        let commandbar = document.querySelector(".ms-commandbar-secondarycommand");

        let wf = commandbar.querySelectorAll('button[name="site workflows"]');
        wf[0].remove();
        let ss = commandbar.querySelectorAll('button[name="site settings"]');
        ss[0].remove();
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
    var userType = await this._checkUser();
    console.log("usertype " + userType)
    // Add page
    if (userType != "owner") {
      var aP = settingsPane.querySelectorAll('a[href="' + this.context.pageContext.web.serverRelativeUrl +'/_layouts/15/CreateSitePage.aspx"]');
      if (aP.length > 0) aP[0].remove();
      aP = settingsPane.querySelectorAll("#SuiteMenu_zz8_MenuItemAddPage");
      if (aP.length > 0) aP[0].remove();
      console.log("not owner");
    }

    //Add app
    var aP = settingsPane.querySelectorAll('a[href="' + this.context.pageContext.web.serverRelativeUrl + '/_layouts/15/appStore.aspx#myApps?entry=SettingAddAnApp"]');
    if (aP.length > 0) aP[0].remove();
    aP = settingsPane.querySelectorAll("#SuiteMenu_zz5_MenuItemCreate");
    if (aP.length > 0) aP[0].remove();

    //Global Navigation
    var gN = settingsPane.querySelectorAll('a[href="javascript:_spLaunchGlobalNavSettings();"]');
    if (gN.length > 0) gN[0].remove();
    gN = settingsPane.querySelectorAll("#GLOBALNAV_SETTINGS_SUITENAVID");
    if (gN.length > 0) gN[0].remove();

    //Hub settings
    var hS = settingsPane.querySelectorAll('a[href="javascript:_spLaunchHubSettings();"]');
    if (hS.length > 0) hS[0].remove();
    hS = settingsPane.querySelectorAll("#SUITENAV_HUB_SETTINGS");
    if (hS.length > 0) hS[0].remove();

    //Site settings
    var sT = settingsPane.querySelectorAll('a[href="' + this.context.pageContext.web.serverRelativeUrl + '/_layouts/15/settings.aspx"]');
    if (sT.length > 0) sT[0].remove();
    sT = settingsPane.querySelectorAll("#SuiteMenu_zz7_MenuItem_Settings");
    if (sT.length > 0) sT[0].remove();

    // Site permissions
    var sP = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSitePermissions();"]');
    if(sP.length > 0) sP[0].remove();
    sP = settingsPane.querySelectorAll("#SUITENAV_SITE_PERMISSIONS");
    if (sP.length > 0) sP[0].remove();

    // Site information
    if (userType === "owner") {
      var sI = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteSettings();"]');
      if (sI.length > 0) {
        let element: HTMLElement = sI[0] as HTMLElement;
        element.onclick = () => {
          window.setTimeout(() => {
            var siteSettingsPane = document.getElementsByClassName("ms-SiteSettingsPanel-SiteInfo");
            if (siteSettingsPane.length > 0) {
              window.setTimeout(() => {
                var jhs = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-joinHubSite");
                if (jhs.length > 0) jhs[0].remove();
              }, 300);
              var c = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-classification");
              if (c.length > 0) c[0].remove();
              var p = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-PrivacyDropdown");
              if (p.length > 0) p[0].remove();
              var ht = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-HelpText");
              if (ht.length > 0) ht[0].remove();
            }
          }, 500);
        }
      }
      console.log("owner")
    } else {
      var sI = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteSettings();"]');
      if (sI.length > 0) sI[0].remove();
      sI = settingsPane.querySelectorAll("#SUITENAV_SITE_INFORMATION");
      if (sI.length > 0) sI[0].remove();
      console.log("not Owner")
    }

    //var sI2 = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SiteInformation");
    //if(sI2.length > 0) sI2[0].remove();
    // Apply Site Template
    var sT = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteTemplates();"]');
    if (sT.length > 0) sT[0].remove();
    sT = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_WebTempaltesGallery");
    if (sT.length > 0) sT[0].remove();

    //Site Performance
    var sP = settingsPane.querySelectorAll('a[href="javascript:_spSitePerformanceScorePage();"]');
    if (sP.length > 0) sP[0].remove();
    sP = settingsPane.querySelectorAll("#SUITENAV_SCORE_PAGE");
    if (sP.length > 0) sP[0].remove();

    // Change the look
    var cTL = settingsPane.querySelectorAll('a[href="javascript:_spLaunchChangeTheLookPanel();"]');
    if(cTL.length > 0) cTL[0].remove();
    cTL = settingsPane.querySelectorAll("#Change_The_Look");
    if (cTL.length > 0) cTL[0].remove();

    // Schedule Site Launch
    var sSL = settingsPane.querySelectorAll('a[href="javascript:_spSiteLaunchSchedulerPage();"]');
    if (sSL.length > 0) sSL[0].remove();
    sSL = settingsPane.querySelectorAll("#SITE_LAUNCH_SUITENAVID");
    if (sSL.length > 0) sSL[0].remove();

    // Site Designs
    var sD = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteDesignProgress();"]');
    if(sD.length > 0) sD[0].remove();
    sD = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SiteDesigns");
    if(sD.length > 0) sD[0].remove();
  }
}
