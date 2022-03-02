import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import { graph } from "@pnp/graph/presets/all";
import "@pnp/graph/users";
import { PermissionKind, stringIsNullOrEmpty } from '@pnp/pnpjs';
import { sp } from "@pnp/sp/presets/all";

const LOG_SOURCE: string = 'WallsApplicationCustomizer';

export interface IWallsApplicationCustomizerProperties {
  adminGroupIds: string;      // The security group GUIDS from AAD that are considered admins
  adminSelectorsCSS: string;  // The selectors for elements we're blocking for admin 
  ownerSelectorsCSS: string;  //                                           for owner
  memberSelectorsCSS: string; //                                           for member and regular
};

// These are security groups defined in azure active directory
// Design, Support, SCA
var adminSecurityGroups = {
  development: 'fae18680-a627-41ed-804a-542dc00531af, e90c926a-e9d0-4f6e-8ccd-3a6938615379, c24ed13a-bbf4-455f-87dd-dff554814df2',
  production: '315f2b29-7a6d-4715-b3cf-3af28d0ddf4b, 24998f56-6911-4041-b4d1-f78452341da6, 77582dc8-6ce7-4a43-9bdd-96e95ce5c78c'
};

enum userType { user, member, owner, admin };

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

      // Check if the group is in the admin groups list. Remove any spaces (should be a list of GUIDS seperated by commas)
      if (this.foundIn(groups.id, this.properties.adminGroupIds.replace(/\s/g, ''))) {
        retVal = userType.admin;
      }
    }

    //If user is an admin, it should keep the admin access not owner
    if (isOwner && retVal != userType.admin) { 
      retVal = userType.owner;
    }

    return retVal;
  }

  // Insert the CSS into the document's head depending on user type
  public addWallsCSS(): void {
    let css: string;

    switch(this.userType) {
      case userType.user:
      case userType.member:
        css = this.createCSS(this.properties.memberSelectorsCSS);
        break;
      case userType.owner:
        css = this.createCSS(this.properties.ownerSelectorsCSS);
        break;
      case userType.admin:
        css = this.createCSS(this.properties.adminSelectorsCSS);
        break;
      default:
        css = '';
        break;
    }

    document.head.insertAdjacentHTML('beforeend', '<style>' + css + '</style>');
  }

  // Go through the list of selectors and generate CSS that hides the elements
  public createCSS(listOfSelectors: string): string {
    if(stringIsNullOrEmpty(listOfSelectors))
      return "";

    let css: string = "";
    const list = listOfSelectors.trim().split(',');

    for(let i = 0; i < list.length; i++) {
      if(list[i] === '') continue;
      css += list[i].trim() + ' { display: none !important } ';
    }
    
    return css.slice(0, -1); // remove trailing space
  }

  public foundIn(identifier: string, commaSeperatedString: string): boolean {
    if(stringIsNullOrEmpty(commaSeperatedString))
      return false;

    var arr = commaSeperatedString.split(',');

    for(let i = 0; i < arr.length; i++) {
      if(identifier == arr[i])
        return true;
    }
    
    return false;
  }
}
