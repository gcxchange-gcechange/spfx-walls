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

// TODO: Add comments for the corresponding groups in AAD
enum userType {
  user,
  member,
  owner,
  admin
}

export default class WallsApplicationCustomizer
  extends BaseApplicationCustomizer<IWallsApplicationCustomizerProperties> {

  private userType: userType;
  private isSettingsOpen: boolean = false;
  private isMobile: boolean = false;

  @override
  public async onInit(): Promise<void> {
    this.userType = await this._checkUser();

    if (this.userType != userType.admin) {
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
    let retVal = userType.user;
    let templateType = this.context.pageContext.web.templateName; // 64: teams, 68: comms

    if (sp.web.hasPermissions(permissions, PermissionKind.ManageWeb) 
    && sp.web.hasPermissions(permissions, PermissionKind.ManagePermissions) 
    && sp.web.hasPermissions(permissions, PermissionKind.CreateGroups)) {

      isOwner = true;  // check if user is a owner by checking the permission
    }

    let user: any[] = await graph.me.memberOf();

    for (let groups of user) {
      if (templateType == "64") { // If site is a teams site (no group member on comms site)
        if (groups.id === this.context.pageContext.site.group.id["_guid"]) { // If user is member of the group
          retVal = userType.member;
        }
      }

      if (groups.id === "c32ff810-25ae-43d3-af87-0b2b5c41dc09" // SCA
        || groups.id === "315f2b29-7a6d-4715-b3cf-3af28d0ddf4b" // UX DESIGN
        || groups.id === "24998f56-6911-4041-b4d1-f78452341da6") { // SUPPORT
          
        retVal = userType.admin;
      }
    }

    //If user is an admin, it should keep the admin access not owner
    if (isOwner && retVal != userType.admin) { 
      retVal = userType.owner;
    }

    return retVal;
  }

  public _render() {

    // Setup the debounced tracking for window resize events
    // We need this to be able to tell if we're in mobile view or not, as the settings buttons are different.
    window.addEventListener('resize', this._debounce(function(){
      this.isMobile = this._isMobile();
    }));

    // Wait for settings button to load, then bind to the click event
    this._awaitSettingsButtonLoad();

    // Site contents page
    if (this.context.pageContext.site.serverRequestPath === "/_layouts/15/viewlsts.aspx") {
      window.setTimeout(() => {
        let commandbar = document.querySelector(".ms-CommandBar-secondaryCommand");
        let wf = commandbar.querySelectorAll('button[name="Site workflows"]');
        wf[0].remove();
        let ss = commandbar.querySelectorAll('button[name="Site settings"]');
        ss[0].remove();
      }, 175);
    }
  }

  // This function will wait for either the mobile or desktop settings button to load onto the page.
  public _awaitSettingsButtonLoad() {
    let interval = setInterval(() => {
      var settingsButton = document.getElementById('O365_MainLink_Settings');

      // Look for desktop layout
      if(settingsButton) {

        // If the user doesn't have any permissions we can hide the settings button.
        if(this.userType === userType.user) {
          settingsButton.style.display = "none";
        }
        
        this._setupEvents(settingsButton);

        clearInterval(interval);
      }
      else {
        // Check for mobile layout
        settingsButton = document.getElementById('O365_MainLink_Affordance');

        if(settingsButton) {

          this.isMobile = true;
          this._setupEvents(settingsButton);

          clearInterval(interval);
        }
      } 
    }, 100);
  }
  
  // This sets up the event listeners 
  public _setupEvents(settingsButton: HTMLElement) {
    var scope = this; // need to pass in scope since we're nesting anonymous functions

    if(!this.isMobile) {

      settingsButton.addEventListener('click', function() {
        scope.isSettingsOpen = !scope.isSettingsOpen;
  
        // Once the settings are opening we can start looking for the pane
        if(scope.isSettingsOpen) {
          scope._awaitSettingsPaneLoad();
        }
      });

      this._setCloseButton('O365_MainLink_Help');
    }
    else {

      settingsButton.addEventListener('click', function () {
        var timeout = 0;

        // Give some delay for the drop down menu to load
        var interval = setInterval(function() {
          var mobileSettings = document.getElementById('O365_MainLink_Settings_Affordance');

          if(mobileSettings) {

            if(scope.userType === userType.user) {
              mobileSettings.parentElement.style.display = "none";
            }
            else {
              mobileSettings.addEventListener('click', function() {

                scope.isSettingsOpen = !scope.isSettingsOpen;
  
                // Once the settings are opening we can start looking for the pane
                if(scope.isSettingsOpen) {
                  scope._awaitSettingsPaneLoad();
                }
              });
            }
            scope._setCloseButton('O365_MainLink_Help_Affordance');
          }

          // Wait up to half a second to find the settings button before clearing the interval
          if(timeout++ >= 50) {
            clearInterval(interval);
          }
        }, 10); 
      });
    }

    // These are the buttons that share open/close state control with the settings pane regardless of mobile layout
    this._setCloseButton('TipsNTricksButton');
    this._setCloseButton('O365_MainLink_Me');
  }


  public _awaitSettingsPaneLoad() {
    var interval = setInterval(() => {
      var settingsPane = document.getElementById('SettingsFlexPane');

      if(settingsPane) {

        this._setCloseButton('flexPaneCloseButton');
        this._addWalls();

        clearInterval(interval);
      }
    }, 5); // Small interval since this will only be called when the pane is in the process of being loaded
  }

  // Track the other buttons that automatically close the settings pane.
  public _setCloseButton(id: string) {
    var button = document.getElementById(id);

    if(button) {
      var scope = this;

      button.addEventListener('click', function() {
        scope.isSettingsOpen = false;
      });
    }
  }

  // Look for the mobile settings button to figure out if we're in the mobile layout or not.
  public _isMobile(): boolean {
    var mobileSettings = document.getElementById('O365_MainLink_Affordance');

    // If the layout has changed we need to setup our events again
    if(mobileSettings && !this.isMobile || !mobileSettings && this.isMobile) {
      this._awaitSettingsButtonLoad();
    }

    return mobileSettings ? true : false;
  }

  // Prevents a function from being called too many times within a given time frame
  public _debounce(func, timeout = 50){
    let timer;
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => { func.apply(this, args); }, timeout);
    };
  }

  public async _addWalls() {
    var settingsPane = document.getElementById('FlexPane_Settings');
    
    // Remove options in the settings pane
    if(settingsPane !== null) {
      // Add page
      if (this.userType && this.userType != userType.owner) {
        var aP = settingsPane.querySelectorAll('a[href="' + this.context.pageContext.web.serverRelativeUrl +'/_layouts/15/CreateSitePage.aspx"]');
        if (aP && aP.length > 0) aP[0].remove();
        aP = settingsPane.querySelectorAll("#SuiteMenu_zz8_MenuItemAddPage");
        if (aP && aP.length > 0) aP[0].remove();
      }

      //Add app
      var aP = settingsPane.querySelectorAll('a[href="' + this.context.pageContext.web.serverRelativeUrl + '/_layouts/15/appStore.aspx#myApps?entry=SettingAddAnApp"]');
      if (aP && aP.length > 0) aP[0].remove();
      aP = settingsPane.querySelectorAll("#SuiteMenu_zz5_MenuItemCreate");
      if (aP && aP.length > 0) aP[0].remove();

      //Global Navigation
      var gN = settingsPane.querySelectorAll('a[href="javascript:_spLaunchGlobalNavSettings();"]');
      if (gN && gN.length > 0) gN[0].remove();
      gN = settingsPane.querySelectorAll("#GLOBALNAV_SETTINGS_SUITENAVID");
      if (gN && gN.length > 0) gN[0].remove();

      //Hub settings
      var hS = settingsPane.querySelectorAll('a[href="javascript:_spLaunchHubSettings();"]');
      if (hS && hS.length > 0) hS[0].remove();
      hS = settingsPane.querySelectorAll("#SUITENAV_HUB_SETTINGS");
      if (hS && hS.length > 0) hS[0].remove();

      //Site settings
      var sT = settingsPane.querySelectorAll('a[href="' + this.context.pageContext.web.serverRelativeUrl + '/_layouts/15/settings.aspx"]');
      if (sT && sT.length > 0) sT[0].remove();
      sT = settingsPane.querySelectorAll("#SuiteMenu_zz7_MenuItem_Settings");
      if (sT && sT.length > 0) sT[0].remove();

      // Site permissions
      var sP = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSitePermissions();"]');
      if(sP && sP.length > 0) sP[0].remove();
      sP = settingsPane.querySelectorAll("#SUITENAV_SITE_PERMISSIONS");
      if (sP && sP.length > 0) sP[0].remove();
      sP = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SitePermissions");
      if (sP && sP.length > 0) sP[0].remove();

      // Site information
      if (this.userType === userType.owner) {
        var sI1 = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteSettings();"]');
        var sI2 = settingsPane.querySelectorAll('#SuiteMenu_MenuItem_SiteInformation'); //For site content page

        //Check if on home page or site content page
        if (sI1 && Object.keys(sI1).length > 0) {
          sI = sI1
        } else if (sI2 && Object.keys(sI2).length > 0) {
          sI = sI2
        }

        if (sI && sI.length > 0) {
          let element: HTMLElement = sI[0] as HTMLElement;
          element.onclick = () => {
            window.setTimeout(() => {
              var siteSettingsPane = document.getElementsByClassName("ms-SiteSettingsPanel-SiteInfo");
              if (siteSettingsPane && siteSettingsPane.length > 0) {
                window.setTimeout(() => {
                  var jhs = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-joinHubSite");
                  if (jhs && jhs.length > 0) jhs[0].remove();
                }, 300);
                var c = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-classification");
                if (c && c.length > 0) c[0].remove();
                var p = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-PrivacyDropdown");
                if (p && p.length > 0) p[0].remove();
                var ht = siteSettingsPane[0].getElementsByClassName("ms-SiteSettingsPanel-HelpText");
                if (ht && ht.length > 0) ht[0].remove();
              }
            }, 500);
          }
        }
      } else {
        var sI = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteSettings();"]');
        if (sI && sI.length > 0) sI[0].remove();
        sI = settingsPane.querySelectorAll("#SUITENAV_SITE_INFORMATION");
        if (sI && sI.length > 0) sI[0].remove();
      }

      //var sI2 = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SiteInformation");
      //if(sI2.length > 0) sI2[0].remove();
      // Apply Site Template
      var sT = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteTemplates();"]');
      if (sT && sT.length > 0) sT[0].remove();
      sT = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_WebTempaltesGallery");
      if (sT && sT.length > 0) sT[0].remove();

      //Site Performance
      var sP = settingsPane.querySelectorAll('a[href="javascript:_spSitePerformanceScorePage();"]');
      if (sP && sP.length > 0) sP[0].remove();
      sP = settingsPane.querySelectorAll("#SUITENAV_SCORE_PAGE");
      if (sP && sP.length > 0) sP[0].remove();

      // Change the look
      var cTL = settingsPane.querySelectorAll('a[href="javascript:_spLaunchChangeTheLookPanel();"]');
      if(cTL && cTL.length > 0) cTL[0].remove();
      cTL = settingsPane.querySelectorAll("#Change_The_Look");
      if (cTL && cTL.length > 0) cTL[0].remove();

      // Schedule Site Launch
      var sSL = settingsPane.querySelectorAll('a[href="javascript:_spSiteLaunchSchedulerPage();"]');
      if (sSL && sSL.length > 0) sSL[0].remove();
      sSL = settingsPane.querySelectorAll("#SITE_LAUNCH_SUITENAVID");
      if (sSL && sSL.length > 0) sSL[0].remove();

      // Site Designs
      var sD = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteDesignProgress();"]');
      if(sD && sD.length > 0) sD[0].remove();
      sD = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SiteDesigns");
      if(sD && sD.length > 0) sD[0].remove();
    }
  }
}
