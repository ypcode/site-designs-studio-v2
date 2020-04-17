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
import { Confirm } from "../common/confirm/Confirm";
import { toJSON } from "../../utils/jsonUtils";
import { ExportServiceKey } from "../../services/export/ExportService";
import { ExportPackage } from "../../helpers/ExportPackage";
import { ExportPackageViewer } from "../exportPackageViewer/ExportPackageViewer";
import { useTraceUpdate } from "../../helpers/hooks";

export interface ISiteScriptEditorProps {
    siteScript: ISiteScript;
}

type ExportType = "json" | "PnPPowershell" | "PnPTemplate" | "o365_PS" | "o365_Bash";

interface ISiteScriptEditorState {
    editingSiteScript: ISiteScript;
    updatedContentFrom: "UI" | "CODE" | null;
    isValidCode: boolean;
    isSaving: boolean;
    isExportUIOpen: boolean;
    currentExportPackage: ExportPackage;
    currentExportType: ExportType;
}

export const SiteScriptEditor = (props: ISiteScriptEditorProps) => {
    useTraceUpdate('SiteScriptEditor', props);
    const [appContext, execute] = useAppContext<IApplicationState, ActionType>();

    // Get service references
    const siteDesignsService = appContext.serviceScope.consume(SiteDesignsServiceKey);
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);
    const exportService = appContext.serviceScope.consume(ExportServiceKey);

    const [state, setState] = useState<ISiteScriptEditorState>({
        editingSiteScript: null,
        currentExportPackage: null,
        currentExportType: "json",
        isExportUIOpen: false,
        isSaving: false,
        isValidCode: true,
        updatedContentFrom: null
    });
    const { editingSiteScript,
        isValidCode,
        updatedContentFrom,
        isExportUIOpen,
        currentExportType,
        currentExportPackage,
        isSaving } = state;

    // Use refs
    const codeEditorRef = useRef<any>();
    const titleFieldRef = useRef<any>();


    const setLoading = (loading: boolean) => {
        execute("SET_LOADING", { loading });
    };
    // Use Effects
    useEffect(() => {
        if (!props.siteScript.Id) {
            setState({ ...state, editingSiteScript: { ...props.siteScript } as ISiteScript });
            if (titleFieldRef.current) {
                titleFieldRef.current.focus();
            }
            return;
        }

        setLoading(true);
        console.debug("Loading site script...", props.siteScript.Id);
        siteDesignsService.getSiteScript(props.siteScript.Id).then(loadedSiteScript => {
            setState({ ...state, editingSiteScript: loadedSiteScript });
            console.debug("Loaded: ", loadedSiteScript);
        }).catch(error => {
            console.error(`The Site Script ${props.siteScript.Id} could not be loaded`, error);
        }).then(() => {
            setLoading(false);
        });
    }, [props.siteScript]);

    const onTitleChanged = (ev: any, title: string) => {
        const newSiteScript = { ...editingSiteScript, Title: title };
        setState({
            ...state,
            editingSiteScript: newSiteScript
        });
    };

    let currentDescription = useRef<string>(editingSiteScript && editingSiteScript.Description);
    const onDescriptionChanging = (ev: any, description: string) => {
        currentDescription.current = description;
    };

    const onDescriptionInputBlur = (ev: any) => {
        setState({
            ...state,
            editingSiteScript: { ...editingSiteScript, Description: currentDescription.current }
        });
    };

    const onVersionChanged = (ev: any, version: string) => {
        const versionInt = parseInt(version);
        if (!isNaN(versionInt)) {
            setState({
                ...state,
                editingSiteScript: { ...editingSiteScript, Version: versionInt }
            });
        }
    };

    const onSiteScriptContentUpdatedFromUI = (updatedContent: ISiteScriptContent) => {
        const newState = {
            ...state,
            updatedContentFrom: 'UI',
            editingSiteScript: { ...editingSiteScript, Content: updatedContent }
        } as ISiteScriptEditorState;
       setState(newState);
    };

    const onSave = async () => {
        setState({ ...state, isSaving: true });
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

        setState({ ...state, isSaving: false });
    };

    const onDelete = async () => {
        if (!await Confirm.show({
            title: `Delete Site Script`,
            message: `Are you sure you want to delete ${(editingSiteScript && editingSiteScript.Title) || "this Site Script"} ?`
        })) {
            return;
        }

        setState({ ...state, isSaving: true });
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
        setState({ ...state, isSaving: false });
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
                setState({
                    ...state,
                    currentExportPackage: exportPackage,
                    currentExportType: exportType,
                    isExportUIOpen: true
                });
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
            setState({
                ...state,
                updatedContentFrom: null
            });
            return;
        }

        codeUpdateTimeoutHandle = setTimeout(() => {
            try {
                if (siteScriptSchemaService.validateSiteScriptJson(updatedCode)) {
                    const updatedScriptContent = JSON.parse(updatedCode) as ISiteScriptContent;
                    setState({
                        ...state,
                        editingSiteScript: { ...editingSiteScript, Content: updatedScriptContent },
                        updatedContentFrom: 'CODE',
                        isValidCode: true
                    });
                } else {
                    setState({
                        ...state,
                        isValidCode: false
                    });
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
    if (!editingSiteScript) {
        return null;
    }
    console.debug("editingSiteScript: ", editingSiteScript);
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
                <DefaultButton text="Cancel" onClick={() => setState({ ...state, isExportUIOpen: false })} /></Stack>}>
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