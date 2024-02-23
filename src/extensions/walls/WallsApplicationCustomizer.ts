import { override } from "@microsoft/decorators";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";

import { graph } from "@pnp/graph/presets/all";
import "@pnp/graph/users";
import { PermissionKind, stringIsNullOrEmpty } from "@pnp/pnpjs";
import { sp } from "@pnp/sp/presets/all";

export interface IWallsApplicationCustomizerProperties {
  adminGroupIds: string; // The security group GUIDS from AAD that are considered admins
  adminSelectorsCSS: string; // The selectors for elements we're blocking for admin
  ownerSelectorsCSS: string; //                                           for owner
  memberSelectorsCSS: string; //                                           for member and regular
  comunicationSiteSelectorsCSS: string //The selectors for elements we're blocking for communicaation sites only
  adminRedirects: string; // The blocked pages for admins
  ownerRedirects: string; //                       owners
  memberRedirects: string; //                       member and regular
  redirectLandingPage: string; // The page users will be redirected to if they go to a blocked page
  logging: string; // Turn logging to the web console on or off ("true" or "false")
}

enum userType {
  user = "user",
  member = "member",
  owner = "owner",
  admin = "admin",
}

export default class WallsApplicationCustomizer extends BaseApplicationCustomizer<IWallsApplicationCustomizerProperties> {
  private userType: userType;

  @override
  public async onInit(): Promise<void> {
    await super.onInit();
    this.context.application.navigatedEvent.add(this, this._initialize);

    return Promise.resolve();
  }

  public async _initialize() {
    if (this.propertiesExist()) {
   //   location.replace(location.href); 
   // location.reload()    
      this.userType = await this._checkUser();
      this.addWallsCSS();
      this.addWallsRedirect();
    } else {
      if (this.properties.logging === "true") {
        console.log("properties Not Exist;");
      }
    }
  }

  public async _checkUser() {
    graph.setup({
      spfxContext: this.context as any,
    });

    const permissions = await sp.web.getCurrentUserEffectivePermissions();
    let isOwner = false;
    let retVal = userType.user;
    const templateType = this.context.pageContext.web.templateName; // 64: teams, 68: comms

    if (
      sp.web.hasPermissions(permissions, PermissionKind.ManageWeb) &&
      sp.web.hasPermissions(permissions, PermissionKind.ManagePermissions) &&
      sp.web.hasPermissions(permissions, PermissionKind.CreateGroups)
    ) {
      isOwner = true; // check if user is a owner by checking the permission
    }

    const userGroups: any[] = await graph.me.memberOf();

    for (let group of userGroups) {
      if (templateType === "64") {
        // If site is a teams site (no group member on comms site)
        if (group.id === this.context.pageContext.site.group.id["_guid"]) {
          // If user is member of the group
          retVal = userType.member;
        }
      }

      // Check if the group is in the admin groups list. Remove any spaces (should be a list of GUIDS seperated by commas)
      if (
        this.foundIn(
          group.id,
          `${this.properties.adminGroupIds}`.replace(/\s/g, "")
        )
      ) {
        retVal = userType.admin;
        break;
      }
    }

    //If user is an admin, it should keep the admin access not owner
    if (isOwner && retVal !== userType.admin) {
      retVal = userType.owner;
    }
    if (this.properties.logging === "true") {
      console.log("User Type", retVal);
    }

    return retVal;
  }

  // Insert the CSS into the document's head depending on user type
  public addWallsCSS(): void {
    let css: string = "";

    switch (this.userType) {
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
    }

    // Add CSS for community sites
    if (this.context.pageContext.web.templateName === "68") {
      // test that I added the previous css and the new string together correctly. 
      // I think just a space and a comma will work but I didn't test.
      css += this.createCSS(this.properties.comunicationSiteSelectorsCSS); 
    }


    console.log("Sensitive group info");
    // let siteHeader = document.querySelector('[class^="actionsWrapper-"]');
    // if (siteHeader.querySelector('[class^="groupInfo-"]')) {
    //   siteHeader
    //     .querySelector<HTMLElement>('[data-automationid="SiteHeaderGroupType"]')
    //     .remove();
    //   const spans = siteHeader.querySelectorAll<HTMLElement>("span");
    //   for (let i = 0; i < spans.length; i++) {
    //     // eslint-disable-next-line eqeqeq
    //     if (spans[i].innerHTML == " | ") {
    //       spans[i].remove();
    //     }
    //   }
    // }

    // Overwrite the styles if they exist
    let existingStyles = document.getElementById('gc-walls-css');
    if (existingStyles) {
      existingStyles.parentNode.removeChild(existingStyles);
    }

    document.head.insertAdjacentHTML("beforeend", '<style id="gc-walls-css">' + css + '</style>');

    if (this.properties.logging === "true") {
      console.log("spfx-walls - Adding CSS for " + this.userType);
      console.log("final css" ,css);
    }
  }

  public addWallsRedirect(): void {
    let blockedPages: any;

    switch (this.userType) {
      case userType.user:
      case userType.member:
        blockedPages = this.properties.memberRedirects;
        break;
      case userType.owner:
        blockedPages = this.properties.ownerRedirects;
        break;
      case userType.admin:
        blockedPages = this.properties.adminRedirects;
        break;
    }

    if (this.properties.logging === "true") {
      console.log("spfx-walls - Adding blocked pages for " + this.userType);
      console.log(blockedPages);
    }

    blockedPages = blockedPages.trim().split(",");

    for (let i = 0; i < blockedPages.length; i++) {
      if (blockedPages[i] === "") continue;

      if (
        window.location.href
          .toLocaleLowerCase()
          .indexOf(blockedPages[i].trim().toLocaleLowerCase()) != -1
      ) {
        if (this.properties.redirectLandingPage != "") {
          window.location.replace(this.properties.redirectLandingPage);
        } else {
          window.location.replace(window.location.origin);
        }
      }
    }
  }

  // Go through the list of selectors and generate CSS that hides the elements
  public createCSS(listOfSelectors: string): string {
    if (stringIsNullOrEmpty(listOfSelectors)) return "";

    let css: string = "";
    const list = listOfSelectors.trim().split(",");

    for (let i = 0; i < list.length; i++) {
      if (list[i] === "") continue;
      css += list[i].trim() + " { display: none !important } ";
      let found =this.foundIn(
        list[i],
        `${this.properties.comunicationSiteSelectorsCSS}`)
      if (
       !found && this.context.pageContext.web.templateName!=="64"
      ){
      this.setRemoveInterval(list[i].trim());
      }
    }

    return css.slice(0, -1); // remove trailing space
  }


  // Setup an interval for each selector to remove the element from the DOM when it's found
  // Defaulted to run every 5 seconds with a 5min timeout if it doesn't find the element.
  public setRemoveInterval(
    selector: string,
    intervalTime: number = 5000,
    timeout: number = 1500000
  ): void {
    if (stringIsNullOrEmpty(selector)) return;

    // eslint-disable-next-line @typescript-eslint/no-this-alias
    let scope = this;
    let interval = setInterval(function () {
      let element = document.querySelector(selector);

      if (element) {
        if (scope.properties.logging === "true") {
          console.log("spfx-walls - Removing element: " + element);
        }

        element.remove();
        clearInterval(interval);
      }

      timeout -= intervalTime;

      if (timeout <= 0) {
        if (this.properties.logging === "true") {
          console.log(
            "spfx-walls - Timeout reached attempting to find: " + selector
          );
        }

        clearInterval(interval);
      }
    }, intervalTime);
  }

  public foundIn(identifier: string, commaSeperatedString: string): boolean {
    if (
      stringIsNullOrEmpty(identifier) ||
      stringIsNullOrEmpty(commaSeperatedString)
    )
      return false;

    let arr = commaSeperatedString.split(",");

    for (let i = 0; i < arr.length; i++) {
      if (identifier == arr[i]) return true;
    }

    return false;
  }

  public propertiesExist(): boolean {
    if (this.properties.logging === "true") {
      console.log(
        "this.properties.adminGroupIds",
        this.properties.adminGroupIds
      );
      console.log(
        "this.properties.adminSelectorsCSS",
        this.properties.adminSelectorsCSS
      );
      console.log("memberSelectorsCSS", this.properties.memberSelectorsCSS);
      console.log("ownerSelectorsCSS", this.properties.ownerSelectorsCSS);
      console.log("logging", this.properties.logging);
      console.log("adminRedirects", this.properties.adminRedirects);
      console.log("ownerRedirects", this.properties.ownerRedirects);
      console.log("memberRedirects", this.properties.memberRedirects);
      console.log("redirectLandingPage", this.properties.redirectLandingPage);
    }
    if (
      this.properties.adminGroupIds === undefined ||
      typeof this.properties.adminGroupIds !== "string" ||
      this.properties.adminSelectorsCSS === undefined ||
      typeof this.properties.adminSelectorsCSS !== "string" ||
      this.properties.memberSelectorsCSS === undefined ||
      typeof this.properties.memberSelectorsCSS !== "string" ||
      this.properties.ownerSelectorsCSS === undefined ||
      typeof this.properties.ownerSelectorsCSS !== "string" ||
      this.properties.logging === undefined ||
      typeof this.properties.logging !== "string" ||
      this.properties.adminRedirects === undefined ||
      typeof this.properties.adminRedirects !== "string" ||
      this.properties.ownerRedirects === undefined ||
      typeof this.properties.ownerRedirects !== "string" ||
      this.properties.memberRedirects === undefined ||
      typeof this.properties.memberRedirects !== "string" ||
      this.properties.redirectLandingPage === undefined ||
      typeof this.properties.redirectLandingPage !== "string"
    ) {
      return false;
    }

    return true;
  }
}
