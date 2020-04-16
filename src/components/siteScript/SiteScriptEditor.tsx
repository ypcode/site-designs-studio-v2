import * as React from "react";
import { useState, useEffect, useRef } from "react";
import styles from "./SiteScriptEditor.module.scss";
import { TextField, PrimaryButton, Label, Stack, DefaultButton, ProgressIndicator, MessageBarType, CommandButton, IContextualMenuProps, Panel, PanelType, Pivot, PivotItem, Icon } from "office-ui-fabric-react";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType, ISetAllAvailableSiteScripts, IGoToActionArgs, ISetUserMessageArgs } from "../../app/IApplicationAction";
import { SiteScriptDesigner } from "./SiteScriptDesigner";
import { SiteDesignsServiceKey } from "../../services/siteDesigns/SiteDesignsService";
import { ISiteScript, ISiteScriptContent } from "../../models/ISiteScript";
import CodeEditor, { monaco } from "@monaco-editor/react";
import { SiteScriptSchemaServiceKey } from "../../services/siteScriptSchema/SiteScriptSchemaService";
import { useConstCallback } from "@uifabric/react-hooks";
import { Confirm } from "../common/confirm/Confirm";
import { toJSON } from "../../utils/jsonUtils";
import { ExportServiceKey } from "../../services/export/ExportService";
import { ExportPackage } from "../../helpers/ExportPackage";
import { ExportPackageViewer } from "../exportPackageViewer/ExportPackageViewer";

export interface ISiteScriptEditorProps {
    siteScript: ISiteScript;
}

type ExportType = "json" | "PnPPowershell" | "PnPTemplate" | "o365_PS" | "o365_Bash";

export const SiteScriptEditor = (props: ISiteScriptEditorProps) => {

    const [appContext, execute] = useAppContext<IApplicationState, ActionType>();

    // Get service references
    const siteDesignsService = appContext.serviceScope.consume(SiteDesignsServiceKey);
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);
    const exportService = appContext.serviceScope.consume(ExportServiceKey);

    // Use state values
    const [editingSiteScript, setEditingSiteScript] = useState<ISiteScript>({ ...(props.siteScript || {} as ISiteScript) });
    const [updatedContentFrom, setUpdatedContentFrom] = useState<"UI" | "CODE" | null>(null);
    const [isValidCode, setIsValidCode] = useState<boolean>(true);
    const [isSaving, setIsSaving] = useState<boolean>(false);
    const [isExportUIOpen, setIsExportUIOpen] = useState<boolean>(false);
    const [currentExportPackage, setCurrentExportPackage] = useState<ExportPackage>(null);
    const [currentExportType, setCurrentExportType] = useState<ExportType>("json");

    // Use refs
    const codeEditorRef = useRef<any>();
    const titleFieldRef = useRef<any>();


    const setLoading = (loading: boolean) => {
        execute("SET_LOADING", { loading });
    };

    // Use Effects
    useEffect(() => {
        if (!props.siteScript.Id) {
            setEditingSiteScript(props.siteScript);
            if (titleFieldRef.current) {
                titleFieldRef.current.focus();
            }
            return;
        }

        setLoading(true);
        console.debug("Loading site script...", props.siteScript.Id);
        siteDesignsService.getSiteScript(props.siteScript.Id).then(loadedSiteScript => {
            setEditingSiteScript(loadedSiteScript);
            console.debug("Loaded: ", loadedSiteScript);
        }).catch(error => {
            console.error(`The Site Script ${props.siteScript.Id} could not be loaded`, error);
        }).then(() => {
            setLoading(false);
        });
    }, [props.siteScript]);

    const onTitleChanged = (ev: any, title: string) => {
        setEditingSiteScript({ ...editingSiteScript, Title: title });
    };

    let currentDescription = useRef<string>(editingSiteScript.Description);
    const onDescriptionChanging = (ev: any, description: string) => {
        currentDescription.current = description;
    };

    const onDescriptionInputBlur = useConstCallback((ev: any) => {
        setEditingSiteScript({ ...editingSiteScript, Description: currentDescription.current });
    });

    const onVersionChanged = (ev: any, version: string) => {
        const versionInt = parseInt(version);
        if (!isNaN(versionInt)) {
            setEditingSiteScript({ ...editingSiteScript, Version: versionInt });
        }
    };

    const onSiteScriptContentUpdatedFromUI = (updatedContent: ISiteScriptContent) => {
        setUpdatedContentFrom("UI");
        setEditingSiteScript({ ...editingSiteScript, Content: updatedContent });
    };

    const onSave = async () => {
        setIsSaving(true);
        try {
            await siteDesignsService.saveSiteScript(editingSiteScript);
            const refreshedSiteScripts = await siteDesignsService.getSiteScripts();
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${editingSiteScript.Title} has been successfully saved.`,
                    messageType: MessageBarType.success
                }
            } as ISetUserMessageArgs);
            execute("SET_ALL_AVAILABLE_SITE_SCRIPTS", { siteScripts: refreshedSiteScripts } as ISetAllAvailableSiteScripts);
            // If it is a brand new script, force redirect to the script list
            if (!editingSiteScript.Id) {
                execute("GO_TO", { page: "SiteScriptsList" } as IGoToActionArgs);
            }
        } catch (error) {
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${editingSiteScript.Title} could not be saved.`,
                    messageType: MessageBarType.error
                }
            } as ISetUserMessageArgs);
            console.error(error);
        }

        setIsSaving(false);
    };

    const onDelete = async () => {
        if (!await Confirm.show({
            title: `Delete Site Script`,
            message: `Are you sure you want to delete ${(editingSiteScript && editingSiteScript.Title) || "this Site Script"} ?`
        })) {
            return;
        }

        setIsSaving(true);
        try {
            await siteDesignsService.deleteSiteScript(editingSiteScript);
            const refreshedSiteScripts = await siteDesignsService.getSiteScripts();
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${editingSiteScript.Title} has been successfully deleted.`,
                    messageType: MessageBarType.success
                }
            } as ISetUserMessageArgs);
            execute("SET_ALL_AVAILABLE_SITE_SCRIPTS", { siteScripts: refreshedSiteScripts } as ISetAllAvailableSiteScripts);
            execute("GO_TO", { page: "SiteScriptsList" } as IGoToActionArgs);
        } catch (error) {
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${editingSiteScript.Title} could not be deleted.`,
                    messageType: MessageBarType.error
                }
            } as ISetUserMessageArgs);
            console.error(error);
        }
        setIsSaving(false);
    };

    const onExportRequested = (exportType?: ExportType) => {
        let exportPromise: Promise<ExportPackage> = null;
        switch (exportType) {
            case "PnPPowershell":
                exportPromise = exportService.generateSiteScriptPnPPowershellExportPackage(editingSiteScript);
                break;
            case "PnPTemplate":
                break; // Not yet supported
            case "o365_PS":
                exportPromise = exportService.generateSiteScriptO365CLIExportPackage(editingSiteScript, "Powershell");
                break;
            case "o365_Bash":
                exportPromise = exportService.generateSiteScriptO365CLIExportPackage(editingSiteScript, "Bash");
                break;
            case "json":
            default:
                exportPromise = exportService.generateSiteScriptJSONExportPackage(editingSiteScript);
                break;
        }

        if (exportPromise) {
            exportPromise.then(exportPackage => {
                setCurrentExportPackage(exportPackage);
                setCurrentExportType(exportType);
                setIsExportUIOpen(true);
            });
        }
    };

    let codeUpdateTimeoutHandle: any = null;
    const onCodeChanged = (updatedCode: string) => {
        if (!updatedCode) {
            return;
        }

        if (codeUpdateTimeoutHandle) {
            clearTimeout(codeUpdateTimeoutHandle);
        }

        if (updatedContentFrom == "UI") {
            // Not trigger the change of state if the script content was updated from UI
            console.debug("The code has been modified after a change in designer. The event will not be propagated");
            setUpdatedContentFrom(null);
            return;
        }

        codeUpdateTimeoutHandle = setTimeout(() => {
            try {
                if (siteScriptSchemaService.validateSiteScriptJson(updatedCode)) {
                    const updatedScriptContent = JSON.parse(updatedCode) as ISiteScriptContent;
                    setEditingSiteScript({ ...editingSiteScript, Content: updatedScriptContent });
                    setUpdatedContentFrom("CODE");
                    setIsValidCode(true);
                } else {
                    setIsValidCode(false);
                }
            } catch (error) {
                console.warn("Code is not valid site script JSON");
            }
        }, 500);
    };

    const editorDidMount = (_, editor) => {

        const schema = siteScriptSchemaService.getSiteScriptSchema();
        codeEditorRef.current = editor;
        monaco.init().then(monacoApi => {
            monacoApi.languages.json.jsonDefaults.setDiagnosticsOptions({
                schemas: [{
                    uri: 'schema.json',
                    schema
                }],

                validate: true,
                allowComments: false
            });
        }).catch(error => {
            console.error("An error occured while trying to configure code editor");
        });


        editor.onDidChangeModelContent(ev => {
            if (codeEditorRef && codeEditorRef.current) {
                onCodeChanged(codeEditorRef.current.getValue());
            }
        });
    };

    const checkIsValidForSave: () => [boolean, string?] = () => {
        if (!editingSiteScript) {
            return [false, "Current Site Script not defined"];
        }

        if (!editingSiteScript.Title) {
            return [false, "Please set the title of the Site Script..."];
        }

        if (!isValidCode) {
            return [false, "Please check the validity of the code..."];
        }

        return [true];
    };

    const isLoading = appContext.isLoading;
    const [isValidForSave, validationMessage] = checkIsValidForSave();
    return <div className={styles.SiteScriptEditor}>
        <div className={styles.row}>
            <div className={styles.columnLayout}>
                <div className={styles.row}>
                    <div className={styles.column11}>
                        <TextField
                            styles={{
                                field: {
                                    fontSize: "32px",
                                    lineHeight: "45px",
                                    height: "45px"
                                },
                                root: {
                                    height: "60px",
                                    marginTop: "5px",
                                    marginBottom: "5px"
                                }
                            }}
                            placeholder="Enter the name of the Site Script..."
                            borderless
                            componentRef={titleFieldRef}
                            value={editingSiteScript.Title}
                            onChange={onTitleChanged} />
                        {isLoading && <ProgressIndicator />}
                    </div>
                    {!isLoading && <div className={`${styles.column1} ${styles.righted}`}>
                        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 15 }}>
                            <CommandButton disabled={isSaving} iconProps={{ iconName: "More" }} menuProps={{
                                items: [
                                    (editingSiteScript.Id && {
                                        key: 'deleteScript',
                                        text: 'Delete',
                                        iconProps: { iconName: 'Delete' },
                                        onClick: onDelete
                                    }),
                                    {
                                        key: 'export',
                                        text: 'Export',
                                        iconProps: { iconName: 'Download' },
                                        onClick: () => onExportRequested(),
                                        disabled: !isValidForSave
                                    }
                                ].filter(i => !!i),
                            } as IContextualMenuProps} />
                            <PrimaryButton disabled={isSaving || !isValidForSave} text="Save" iconProps={{ iconName: "Save" }} onClick={() => onSave()} />
                        </Stack>
                    </div>}
                </div>
                <div className={styles.row}>
                    {editingSiteScript.Id && <div className={styles.half}>
                        <div className={styles.row}>
                            <div className={styles.column8}>
                                <TextField
                                    label="Id"
                                    readOnly
                                    value={editingSiteScript.Id} />
                            </div>
                            <div className={styles.column4}>
                                <TextField
                                    label="Version"
                                    value={editingSiteScript.Version.toString()}
                                    onChange={onVersionChanged} />
                            </div>
                        </div>
                    </div>}
                    <div className={styles.half}>
                        <TextField
                            label="Description"
                            value={editingSiteScript.Description}
                            multiline={true}
                            rows={2}
                            borderless
                            placeholder="Enter a description for the Site Script..."
                            onChange={onDescriptionChanging}
                            onBlur={onDescriptionInputBlur}
                        />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <Label>Actions</Label>
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.designerWorkspace}>
                        <SiteScriptDesigner
                            siteScriptContent={editingSiteScript.Content}
                            onSiteScriptContentUpdated={onSiteScriptContentUpdatedFromUI} />
                    </div>
                    <div className={styles.codeEditorWorkspace}>
                        <CodeEditor
                            height="80vh"
                            language="json"
                            options={{
                                folding: true,
                                renderIndentGuides: true,
                                minimap: {
                                    enabled: false
                                }
                            }}
                            value={toJSON(editingSiteScript.Content)}
                            editorDidMount={editorDidMount}
                        />
                    </div>
                </div>
            </div>
        </div>
        <Panel isOpen={isExportUIOpen}
            type={PanelType.large}
            headerText="Export Site Script"
            onRenderFooterContent={(p) => <Stack horizontalAlign="end" horizontal tokens={{ childrenGap: 10 }}>
                <PrimaryButton iconProps={{ iconName: "Download" }} text="Download" onClick={() => currentExportPackage && currentExportPackage.download()} />
                <DefaultButton text="Cancel" onClick={() => setIsExportUIOpen(false)} /></Stack>}>
            <Pivot
                selectedKey={currentExportType}
                onLinkClick={(item) => onExportRequested(item.props.itemKey as ExportType)}
                headersOnly={true}
            >
                <PivotItem headerText="JSON" itemKey="json" />
                <PivotItem headerText="PnP Powershell" itemKey="PnPPowershell" />
                {/* <PivotItem headerText="PnP Template" itemKey="PnPTemplate" /> */}
                <PivotItem headerText="O365 CLI (Powershell)" itemKey="o365_PS" />
                <PivotItem headerText="O365 CLI (Bash)" itemKey="o365_Bash" />
            </Pivot>
            {currentExportPackage && <ExportPackageViewer exportPackage={currentExportPackage} />}
        </Panel>
    </div>;
};