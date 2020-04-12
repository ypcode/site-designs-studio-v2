import * as React from "react";
import { useState, useEffect, useRef } from "react";
import styles from "./SiteScriptEditor.module.scss";
import { TextField, PrimaryButton, Label, Spinner, SpinnerType, Slider, Stack, DefaultButton, ProgressIndicator } from "office-ui-fabric-react";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType, ISetAllAvailableSiteScripts, IGoToActionArgs } from "../../app/IApplicationAction";
import { SiteScriptDesigner } from "./SiteScriptDesigner";
import { SiteDesignsServiceKey } from "../../services/siteDesigns/SiteDesignsService";
import { ISiteScript, ISiteScriptContent, ISiteScriptAction } from "../../models/ISiteScript";
import CodeEditor, { monaco } from "@monaco-editor/react";
import { SiteScriptSchemaServiceKey } from "../../services/siteScriptSchema/SiteScriptSchemaService";
import { useConstCallback } from "@uifabric/react-hooks";
import { Confirm } from "../common/Confirm/Confirm";

export interface ISiteScriptEditorProps {
    siteScript: ISiteScript;
}

const siteScriptContentToCodeString = (siteScriptContent: ISiteScriptContent) => JSON.stringify(siteScriptContent, null, 2);

export const SiteScriptEditor = (props: ISiteScriptEditorProps) => {

    const [appContext, execute] = useAppContext<IApplicationState, ActionType>();

    // Get service references
    const siteDesignsService = appContext.serviceScope.consume(SiteDesignsServiceKey);
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);

    // Use state values
    const [editingSiteScriptContent, setEditingSiteScriptContent] = useState<ISiteScriptContent>(props.siteScript.Content);
    const [isSaving, setIsSaving] = useState<boolean>(false);

    // Use refs
    const codeEditorRef = useRef<any>();
    const titleFieldRef = useRef<any>();
    // We don't use editingSiteScript object as state but as a referenced object
    const editingSiteScriptRef = useRef<ISiteScript>(props.siteScript);
    const getEditingSiteScript = () => editingSiteScriptRef && editingSiteScriptRef.current;
    const setEditingSiteScript = (siteScript: ISiteScript) => {
        if (editingSiteScriptRef && editingSiteScriptRef.current) {
            editingSiteScriptRef.current = siteScript;
        }
    };

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
        console.log("Loading site script...", props.siteScript.Id);
        siteDesignsService.getSiteScript(props.siteScript.Id).then(loadedSiteScript => {
            setEditingSiteScript(loadedSiteScript);
            setEditingSiteScriptContent(loadedSiteScript.Content);
            console.log("Loaded: ", loadedSiteScript);
        }).catch(error => {
            console.error(`The Site Script ${props.siteScript.Id} could not be loaded`, error);
        }).then(() => {
            setLoading(false);
        });
    }, [props.siteScript]);

    const onTitleChanged = (ev: any, title: string) => {
        setEditingSiteScript({ ...getEditingSiteScript(), Title: title });
    };

    let currentDescription = useRef<string>(getEditingSiteScript().Description);
    const onDescriptionChanging = (ev: any, description: string) => {
        currentDescription.current = description;
    };

    const onDescriptionInputBlur = useConstCallback((ev: any) => {
        setEditingSiteScript({ ...getEditingSiteScript(), Description: currentDescription.current });
    });

    const onVersionChanged = (ev: any, version: string) => {
        const versionInt = parseInt(version);
        if (!isNaN(versionInt)) {
            setEditingSiteScript({ ...getEditingSiteScript(), Version: versionInt });
        }
    };

    const onSiteScriptContentUpdated = (updatedContent: ISiteScriptContent) => {
        setEditingSiteScriptContent(updatedContent);
    };

    const onSave = async () => {
        setIsSaving(true);
        const toSave = { ...getEditingSiteScript(), Content: editingSiteScriptContent };
        await siteDesignsService.saveSiteScript(toSave);
        const refreshedSiteScripts = await siteDesignsService.getSiteScripts();
        execute("SET_ALL_AVAILABLE_SITE_SCRIPTS", { siteScripts: refreshedSiteScripts } as ISetAllAvailableSiteScripts);
        setIsSaving(false);
    };

    const onDelete = async () => {
        if (!await Confirm.show({
            title: `Delete Site Script`,
            message: `Are you sure you want to delete ${(editingSiteScriptRef.current && editingSiteScriptRef.current.Title) || "this Site Script"} ?`
        })) {
            return;
        }

        setIsSaving(true);
        await siteDesignsService.deleteSiteScript(editingSiteScriptRef.current);
        const refreshedSiteScripts = await siteDesignsService.getSiteScripts();
        execute("SET_ALL_AVAILABLE_SITE_SCRIPTS", { siteScripts: refreshedSiteScripts } as ISetAllAvailableSiteScripts);
        execute("GO_TO", { page: "SiteScriptsList" } as IGoToActionArgs);
        setIsSaving(false);
    };

    let codeUpdateTimeoutHandle: number = null;
    const onCodeChanged = (updatedCode: string) => {
        if (!updatedCode) {
            return;
        }

        if (codeUpdateTimeoutHandle) {
            clearTimeout(codeUpdateTimeoutHandle);
        }

        codeUpdateTimeoutHandle = setTimeout(() => {
            try {
                const updatedScriptContent = JSON.parse(updatedCode);
                if (siteScriptSchemaService.validateSiteScriptJson(updatedScriptContent)) {
                    setEditingSiteScriptContent(updatedScriptContent);
                }
            } catch (error) {
                console.warn("Code is not valid site script JSON");
            }
        }, 2000);
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

    const isLoading = appContext.isLoading;

    console.log("RERENDER SiteScriptEditor");
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
                            value={getEditingSiteScript().Title}
                            onChange={onTitleChanged} />
                        {isLoading && <ProgressIndicator />}
                    </div>
                    {!isLoading && <div className={`${styles.column1} ${styles.righted}`}>
                        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 15 }}>
                            <DefaultButton disabled={isSaving} text="Delete" iconProps={{ iconName: "Delete" }} onClick={() => onDelete()} />
                            <PrimaryButton disabled={isSaving} text="Save" iconProps={{ iconName: "Save" }} onClick={() => onSave()} />
                        </Stack>
                    </div>}
                </div>
                <div className={styles.row}>
                    {getEditingSiteScript().Id && <div className={styles.half}>
                        <div className={styles.row}>
                            <div className={styles.column8}>
                                <TextField
                                    label="Id"
                                    readOnly
                                    value={getEditingSiteScript().Id} />
                            </div>
                            <div className={styles.column4}>
                                <TextField
                                    label="Version"
                                    value={getEditingSiteScript().Version.toString()}
                                    onChange={onVersionChanged} />
                            </div>
                        </div>
                    </div>}
                    <div className={styles.half}>
                        <TextField
                            label="Description"
                            value={getEditingSiteScript().Description}
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
                            siteScriptContent={editingSiteScriptContent}
                            onSiteScriptContentUpdated={onSiteScriptContentUpdated} />
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
                            value={siteScriptContentToCodeString(editingSiteScriptContent)}
                            editorDidMount={editorDidMount}
                        />
                    </div>
                </div>
            </div>
        </div>
    </div>;
};