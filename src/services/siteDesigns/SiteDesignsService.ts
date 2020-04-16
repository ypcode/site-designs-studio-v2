import { ISiteScript, ISiteScriptContent } from '../../models/ISiteScript';
import { ISiteDesign, WebTemplate, ISiteDesignWithGrantedRights } from '../../models/ISiteDesign';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { assign } from '@microsoft/sp-lodash-subset';
import { ISiteScriptSchemaService, SiteScriptSchemaServiceKey } from '../siteScriptSchema/SiteScriptSchemaService';

export interface IGetSiteScriptFromWebOptions {
	includeBranding?: boolean;
	includeLists?: string[];
	includeRegionalSettings?: boolean;
	includeSiteExternalSharingCapability?: boolean;
	includeTheme?: boolean;
	includeLinksToExportedItems?: boolean;
}

export interface IGetSiteScriptFromExistingResourceResult {
	JSON: ISiteScriptContent;
	Warnings: string[];
}

export interface ISiteDesignsService {
	baseUrl: string;
	getSiteDesigns(): Promise<ISiteDesign[]>;
	getSiteDesign(siteDesignId: string): Promise<ISiteDesign>;
	saveSiteDesign(siteDesign: ISiteDesign): Promise<void>;
	deleteSiteDesign(siteDesign: ISiteDesign | string): Promise<void>;
	getSiteScripts(): Promise<ISiteScript[]>;
	getSiteScript(siteScriptId: string): Promise<ISiteScript>;
	saveSiteScript(siteScript: ISiteScript): Promise<void>;
	deleteSiteScript(siteScript: ISiteScript | string): Promise<void>;
	applySiteDesign(siteDesignId: string, webUrl: string): Promise<void>;
	getSiteScriptFromList(listUrl: string): Promise<IGetSiteScriptFromExistingResourceResult>;
	getSiteScriptFromWeb(webUrl: string, options?: IGetSiteScriptFromWebOptions): Promise<IGetSiteScriptFromExistingResourceResult>;
}

export class SiteDesignsService implements ISiteDesignsService {
	private spHttpClient: SPHttpClient;
	private schemaService: ISiteScriptSchemaService;

	constructor(serviceScope: ServiceScope) {
		serviceScope.whenFinished(() => {
			this.spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
			this.schemaService = serviceScope.consume(SiteScriptSchemaServiceKey);
		});
	}

	public baseUrl: string = '/';

	private _getEffectiveUrl(relativeUrl: string): string {
		return (this.baseUrl + '//' + relativeUrl).replace(/\/{2,}/, '/');
	}

	private _restRequest(url: string, params: any = null): Promise<any> {
		const restUrl = this._getEffectiveUrl(url);
		const options: ISPHttpClientOptions = {
			body: JSON.stringify(params),
			headers: {
				'Content-Type': 'application/json;charset=utf-8',
				ACCEPT: 'application/json; odata.metadata=minimal',
				'ODATA-VERSION': '4.0'
			}
		};
		return this.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, options).then((response) => {
			if (response.status == 204) {
				return {};
			} else {
				return response.json();
			}
		});
	}

	public getSiteDesigns(): Promise<ISiteDesign[]> {
		return this._restRequest(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns'
		).then((resp) => resp.value as ISiteDesign[]);
	}
	public getSiteDesign(siteDesignId: string): Promise<ISiteDesign> {
		return this._restRequest(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata',
			{ id: siteDesignId }
		).then((resp) => resp as ISiteDesign);
	}
	public deleteSiteDesign(siteDesign: ISiteDesign | string): Promise<void> {
		let id = typeof siteDesign === 'string' ? siteDesign as string : (siteDesign as ISiteDesign).Id;
		return this._restRequest(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteDesign',
			{ id: id }
		);
	}
	public saveSiteDesign(siteDesign: ISiteDesign): Promise<void> {
		if (siteDesign.Id) {
			// Update
			return this._restRequest(
				'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign',
				{
					updateInfo: {
						Id: siteDesign.Id,
						Title: siteDesign.Title,
						Description: siteDesign.Description,
						SiteScriptIds: siteDesign.SiteScriptIds,
						WebTemplate: siteDesign.WebTemplate.toString(),
						PreviewImageUrl: siteDesign.PreviewImageUrl,
						PreviewImageAltText: siteDesign.PreviewImageAltText,
						Version: siteDesign.Version,
						IsDefault: siteDesign.IsDefault
					}
				}
			).then(() => {
				const withGrantedRights = (siteDesign as ISiteDesignWithGrantedRights);
				if (withGrantedRights.grantedRightsPrincipals) {
					return this._setSiteDesignRights(siteDesign.Id, withGrantedRights.grantedRightsPrincipals);
				} else {
					return Promise.resolve();
				}
			});
		} else {
			// Create
			return this._restRequest(
				'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign',
				{
					info: {
						Title: siteDesign.Title,
						Description: siteDesign.Description,
						SiteScriptIds: siteDesign.SiteScriptIds,
						WebTemplate: siteDesign.WebTemplate.toString(),
						PreviewImageUrl: siteDesign.PreviewImageUrl,
						PreviewImageAltText: siteDesign.PreviewImageAltText
					}
				}
			);
		}
	}

	private _setSiteDesignRights(siteDesignId: string, principalNames: string[]): Promise<void> {
		// Get the current rights of the site design
		return this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights", { id: siteDesignId })
			.then(existingRights => {
				const existingRightPrincipalNames: string[] = existingRights.value.map(r => r.PrincipalName);
				// Remove the ones not included in specified principalNames
				const toRevokePrincipalNames: string[] = existingRightPrincipalNames.filter(r => principalNames.indexOf(r) < 0);
				const revokePromise = this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights",
					{
						id: siteDesignId,
						principalNames: toRevokePrincipalNames
					});

				// Add the ones from principalNames not included in existing
				const toGrantPrincipalNames: string[] = principalNames.filter(pn => existingRightPrincipalNames.indexOf(pn) < 0);
				const grantPromise = this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights",
					{
						id: siteDesignId,
						principalNames: toRevokePrincipalNames,
						grantedRights: 1, // Means "View" , only supported value currently
					});

				return Promise.all([revokePromise, grantPromise]).then(() => { });
			});
	}

	public getSiteScripts(): Promise<ISiteScript[]> {
		return this._restRequest(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts'
		).then((resp) => resp.value as ISiteScript[]);
	}

	public getSiteScript(siteScriptId: string): Promise<ISiteScript> {
		return this._restRequest(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata',
			{ id: siteScriptId }
		).then((resp) => {
			let siteScript = resp as ISiteScript;
			siteScript.Content = JSON.parse(siteScript.Content as any);
			return siteScript;
		});
	}

	public saveSiteScript(siteScript: ISiteScript): Promise<void> {
		if (siteScript.Id) {
			return this._restRequest(
				'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript',
				{
					updateInfo: {
						Id: siteScript.Id,
						Title: siteScript.Title,
						Description: siteScript.Description,
						Version: siteScript.Version,
						Content: JSON.stringify(siteScript.Content)
					}
				}
			);
		} else {
			return this._restRequest(
				`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title)?@title='${siteScript.Title}'`,
				siteScript.Content
			);
		}
	}
	public deleteSiteScript(siteScript: ISiteScript | string): Promise<void> {
		let id = typeof siteScript === 'string' ? siteScript as string : (siteScript as ISiteScript).Id;
		return this._restRequest(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript',
			{ id: id }
		);
	}

	public applySiteDesign(siteDesignId: string, webUrl: string): Promise<void> {
		// TODO Implement
		return null;
	}
	public getSiteScriptFromList(listUrl: string): Promise<IGetSiteScriptFromExistingResourceResult> {
		return this._restRequest('/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptFromList', {
			listUrl
		}).then(result => {
			console.log("result is : ", result);
			const defaultContent = this.schemaService.getNewSiteScript();
			const siteScriptContent = assign(defaultContent, JSON.parse(result.value)) as ISiteScriptContent;
			siteScriptContent.$schema = "schema.json";
			return { Warnings: [], JSON: siteScriptContent };
		});
	}
	public getSiteScriptFromWeb(webUrl: string, options?: IGetSiteScriptFromWebOptions): Promise<IGetSiteScriptFromExistingResourceResult> {
		const info = {};
		if (options) {
			if (options.includeBranding === true || options.includeBranding === false) {
				info["IncludeBranding"] = options.includeBranding;
			}
			if (options.includeLists !== null || typeof options.includeLists !== "undefined") {
				info["IncludedLists"] = options.includeLists;
			}
			if (options.includeRegionalSettings === true || options.includeRegionalSettings === false) {
				info["IncludeRegionalSettings"] = options.includeRegionalSettings;
			}
			if (options.includeSiteExternalSharingCapability === true || options.includeSiteExternalSharingCapability === false) {
				info["IncludeSiteExternalSharingCapability"] = options.includeSiteExternalSharingCapability;
			}
			if (options.includeTheme === true || options.includeTheme === false) {
				info["IncludeTheme"] = options.includeTheme;
			}
			if (options.includeLinksToExportedItems === true || options.includeLinksToExportedItems === false) {
				info["IncludeLinksToExportedItems"] = options.includeLinksToExportedItems;
			}
		}
		return this._restRequest('/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptFromWeb', {
			webUrl,
			info
		}).then(result => {
			const defaultContent = this.schemaService.getNewSiteScript();
			const siteScriptContent = assign(defaultContent, JSON.parse(result.JSON)) as ISiteScriptContent;
			siteScriptContent.$schema = "schema.json";
			return { Warnings: result.Warnings, JSON: siteScriptContent };
		});
	}
}

export const SiteDesignsServiceKey = ServiceKey.create<ISiteDesignsService>(
	'YPCODE:SDSv2:SiteDesignsService',
	SiteDesignsService
);
