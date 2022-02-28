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

// These are security groups defined in azure active directory
const securityGroups = {
  development: {
    design: 'fae18680-a627-41ed-804a-542dc00531af',
    support: 'e90c926a-e9d0-4f6e-8ccd-3a6938615379',
    sca : 'c24ed13a-bbf4-455f-87dd-dff554814df2',
  },
  production: {
    design: '315f2b29-7a6d-4715-b3cf-3af28d0ddf4b',
    support: '24998f56-6911-4041-b4d1-f78452341da6',
    sca: 'c32ff810-25ae-43d3-af87-0b2b5c41dc09',
  }
}

enum userType {
  user, 
  member,
  owner,
  admin
}

export default class WallsApplicationCustomizer
  extends BaseApplicationCustomizer<IWallsApplicationCustomizerProperties> {

  private userType: userType;

  @override
  public async onInit(): Promise<void> {
    this.userType = await this._checkUser();

    this.addWallsCSS();

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

      /*
        IMPORTANT: Change these to either development or production
                    depending on where you're deploying this extension.
      */
      if (groups.id === securityGroups.development.sca
        || groups.id === securityGroups.development.design
        || groups.id === securityGroups.development.support) {
          
        retVal = userType.admin;
      }
    }

    //If user is an admin, it should keep the admin access not owner
    if (isOwner && retVal != userType.admin) { 
      retVal = userType.owner;
    }

    return retVal;
  }

  public addWallsCSS(): void {
    let css: string;

    switch(this.userType) {
      case userType.user:
      case userType.member:
        css = '#O365_MainLink_Settings { display: none !important; } #O365_MainLink_Affordance { display: none !important; } #FlexPane_Settings { display: none !important; }';
        break;
      case userType.owner:
        css = '#SuiteMenu_zz5_MenuItemCreate { display: none !important; } #GLOBALNAV_SETTINGS_SUITENAVID { display: none !important; } #SUITENAV_HUB_SETTINGS { display: none !important; } .ms-SiteSettingsPanel-HelpText a[href*="layouts/15/settings.aspx"] { display: none !important; } .ms-SiteSettingsPanel-classification { display: none !important; } #SUITENAV_SITE_PERMISSIONS { display: none !important; } #SuiteMenu_MenuItem_WebTempaltesGallery { display: none !important; } #SUITENAV_SCORE_PAGE { display: none !important; } #SITE_LAUNCH_SUITENAVID { display: none !important; } #CHANGE_THE_LOOK { display: none !important; }';
        break;
      case userType.admin:
      default:
        css = '';
        break;
    }

    document.head.insertAdjacentHTML('beforeend', '<style>' + css + '</style>');
  }
}
