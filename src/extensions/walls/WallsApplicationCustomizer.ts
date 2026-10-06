
import { override } from "@microsoft/decorators";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";

import { GraphFI } from "@pnp/graph";
import "@pnp/graph/users";
import { stringIsNullOrEmpty } from "@pnp/core";
import { PermissionKind } from "@pnp/sp/security";

import { spfi, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/security";
import "@pnp/sp/site-users/web";
import * as strings from 'WallsApplicationCustomizerStrings';

export interface IWallsApplicationCustomizerProperties {
  adminGroupIds: string; // The security group GUIDS from AAD that are considered admins
  adminSelectorsCSS: string; // The selectors for elements we're blocking for admin
  ownerSelectorsCSS: string; //                                           for owner
  memberSelectorsCSS: string; //                                           for member and regular
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
    this.context.application.navigatedEvent.add(this, this._removeAppbutton);
    this.context.application.navigatedEvent.add(this, this._removeAIAgentLink);
    this.context.placeholderProvider.changedEvent.add(this, this._removeAIAgentLink);


    return Promise.resolve();
  }

  public async _initialize() {
    if (this.propertiesExist()) {
      this.userType = await this._checkUser();
      this.addWallsCSS();
      this.addWallsRedirect();
    } else {
      if (this.properties.logging === "true") {
        console.log("properties Not Exist;");
      }
    }
  }

  public _removeAIAgentLink() {

    const aiAgentAriaLabel = strings.AIAgentLinkAriaLabel;
    const webPartToolboxAriaLabel = strings.WebPartToolboxAriaLabel;

    const findAndRemoveAIAgent = (): boolean => {

      let removed = false;

      const toolbox = document.querySelector('[data-automationid="SPContentPanelView-container"]');
     // console.log("Toolbox found:", toolbox);

      // check the toolbox panel 
      if (toolbox) {
        const aiAgentLink = toolbox.querySelector(`[aria-label="${aiAgentAriaLabel}"]`);
        //console.log("AI Agent link found in toolbox:", aiAgentLink);

        if (aiAgentLink) {
          // console.log("AI Agent found in toolbox:", aiAgentLink);
          aiAgentLink.remove();
          removed = true;
        }

      }

      

      //check for the main content webpart toolbox callout

      const webPartToolbox = document.querySelector( `[aria-label="${webPartToolboxAriaLabel}"]` ); 

      if (webPartToolbox) { 
        const webpartAIAgentLink = webPartToolbox.querySelector( `[aria-label="${aiAgentAriaLabel}"]` ); 
      
        if (webpartAIAgentLink) { 
          console.log("AI Agent found in web part toolbox - removing"); 
          webpartAIAgentLink.remove(); 
          removed = true; 
        } 
      }


      return removed;
    };

    // Check immediately in case everything is already loaded
    if (findAndRemoveAIAgent()) {
      return;
    }

    // Watch for the toolbox AND its contents to load
    const observer = new MutationObserver(() => {

      if (findAndRemoveAIAgent()) {
        observer.disconnect();

        console.log("AI Agent removed and observer disconnected.");
      }
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true
    });

  }


  public _removeAppbutton() {
    window.addEventListener('click', (event) => {
 
      const targetElement = event.target as HTMLElement;

      if (targetElement.outerText === "New" || targetElement.tagName === "svg") {
          const newButtonChildren =  document.querySelector('[data-automation-id="CommandBarNewDashboardButton"]');
          const previousSibling = newButtonChildren?.previousElementSibling;
    
          if (previousSibling) {
            previousSibling.remove();
          }
      }
      
    })
   
  }


  public async _checkUser() {
    const sp = spfi().using(SPFx(this.context as any));
    const graph = new GraphFI().using(SPFx(this.context as any));

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

    console.log("Sensitive group info");
    let siteHeader = document.querySelector('[class^="actionsWrapper-"]');
    if (siteHeader.querySelector('[class^="groupInfo-"]')) {
      siteHeader
        .querySelector<HTMLElement>('[data-automationid="SiteHeaderGroupType"]')
        .remove();
      const spans = siteHeader.querySelectorAll<HTMLElement>("span");
      for (let i = 0; i < spans.length; i++) {
        // eslint-disable-next-line eqeqeq
        if (spans[i].innerHTML == " | ") {
          spans[i].remove();
        }
      }
    }

    document.head.insertAdjacentHTML("beforeend", "<style>" + css + "</style>");

    if (this.properties.logging === "true") {
      console.log("spfx-walls - Adding CSS for " + this.userType);
      console.log(css);
    }
  }

  public addWallsRedirect(): void {
    let blockedPages:any;

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
      this.setRemoveInterval(list[i].trim());
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
   // let scope = this;
    let interval = setInterval(function () {
      let element = document.querySelector(selector);

      if (element) {
        if (this.properties.logging === "true") {
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

    console.log("this.properties.adminGroupIds", this.properties.adminGroupIds);
    console.log("this.properties.adminSelectorsCSS", this.properties.adminSelectorsCSS);
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
