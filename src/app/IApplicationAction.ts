import { Page } from "./ApplicationState";
import { ISiteDesign } from "../models/ISiteDesign";
import { ISiteScript } from "../models/ISiteScript";

export type ActionType = "GO_TO"
    | "EDIT_SITE_DESIGN"
    | "EDIT_SITE_SCRIPT"
    | "SET_ALL_AVAILABLE_SITE_DESIGNS"
    | "SET_ALL_AVAILABLE_SITE_SCRIPTS"
    | "SET_LOADING";



export interface IGoToActionArgs {
    page: Page;
}

export interface IEditSiteDesignActionArgs {
    siteDesign: ISiteDesign;
}

export interface IEditSiteScriptActionArgs {
    siteScript: ISiteScript;
}

export interface ISetAllAvailableSiteDesigns {
    siteDesigns: ISiteDesign[];
}

export interface ISetAllAvailableSiteScripts {
    siteScripts: ISiteScript[];
}

export interface ISetLoadingArgs {
    loading: boolean;
}