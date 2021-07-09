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
      this._addWalls();
      this.context.application.navigatedEvent.add(this, this._addWalls);
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
      if(groups.roleTemplateId && groups.roleTemplateId === "62e90394-69f5-4237-9190-012177145e10") { // Company
        isAdmin = true;
      } else if(groups.roleTemplateId && groups.roleTemplateId === "f28a1f50-f6e7-4571-818b-6a12f2af6b6c") { // Sharepoint
        isAdmin = true;
      } else if(groups.id === "315f2b29-7a6d-4715-b3cf-3af28d0ddf4b") { // UX DESIGN
        isAdmin = true;
      }
    }

    return isAdmin;
  }

  public async _addWalls() {

      // Site Settings Pane
      // Interval to wait for load of site settings
      const interval = window.setInterval(() => {
        var cog = document.getElementById('O365_MainLink_Settings');
        if (cog) {

          cog.onclick = () => {
            const timeout = window.setTimeout(() => {

              var settingsPane = document.getElementById('FlexPane_Settings');

              if(settingsPane) {
                // Remove options in settings
                // Site permissions
                var sP = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSitePermissions();"]');
                if(sP.length > 0) sP[0].remove();
                sP = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SitePermissions");
                if(sP.length > 0) sP[0].remove();

                // Site information
                var sI = settingsPane.querySelectorAll('a[href="javascript:_spLaunchSiteSettings();"]');
                if(sI.length > 0) sI[0].remove();
                sI = settingsPane.querySelectorAll("#SuiteMenu_MenuItem_SiteInformation");
                if(sI.length > 0) sI[0].remove();

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

            }, 200);
          };

          // No more searching
          window.clearInterval(interval);
        }
      }, 300);

      // Site contents page
      if (this.context.pageContext.site.serverRequestPath === "/_layouts/15/viewlsts.aspx") {
          window.setTimeout(() => {
            let commandBar = document.querySelector(".ms-CommandBar-secondaryCommand");

            let wF = commandBar.querySelectorAll('button[name="Site workflows"]');
            wF[0].remove();
            let sS = commandBar.querySelectorAll('button[name="Site settings"]');
            sS[0].remove();
          }, 100);
      }


  }
}
