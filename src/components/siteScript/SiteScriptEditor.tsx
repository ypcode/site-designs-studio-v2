import * as React from "react";
import { useState, useEffect, useRef } from "react";
import styles from "./SiteScriptEditor.module.scss";
import { TextField, PrimaryButton, Label, Spinner, SpinnerType, Slider } from "office-ui-fabric-react";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType, ISetAllAvailableSiteScripts } from "../../app/IApplicationAction";
import { SiteScriptDesigner } from "./SiteScriptDesigner";
import { SiteDesignsServiceKey } from "../../services/siteDesigns/SiteDesignsService";
import { ISiteScript, ISiteScriptContent, ISiteScriptAction } from "../../models/ISiteScript";
import CodeEditor, { monaco } from "@monaco-editor/react";
import { SiteScriptSchemaServiceKey } from "../../services/siteScriptSchema/SiteScriptSchemaService";

export interface ISiteScriptEditorProps {
    siteScript: ISiteScript;
}

const siteScriptContentToCodeString = (siteScriptContent: ISiteScriptContent) => JSON.stringify(siteScriptContent, null, 4);

export const SiteScriptEditor = (props: ISiteScriptEditorProps) => {

    const [appContext, execute] = useAppContext<IApplicationState, ActionType>();

    // Get service references
    const siteDesignsService = appContext.serviceScope.consume(SiteDesignsServiceKey);
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);

    // Use state values
    const [editingSiteScript, setEditingSiteScript] = useState<ISiteScript>(props.siteScript);
    const [focusedField, setFocusedField] = useState<string>("");
    const [isSaving, setIsSaving] = useState<boolean>(false);


    // Use refs
    const codeEditorRef = useRef<any>();

    const setLoading = (loading: boolean) => {
        execute("SET_LOADING", { loading });
    };

    // Use Effects
    useEffect(() => {
        if (!props.siteScript.Id) {
            setEditingSiteScript(props.siteScript);
            return;
        }

        setLoading(true);
        console.log("Loading site script...", props.siteScript.Id);
        siteDesignsService.getSiteScript(props.siteScript.Id).then(loadedSiteScript => {
            setEditingSiteScript(loadedSiteScript);
            console.log("Loaded: ", loadedSiteScript);
        }).catch(error => {
            console.error(`The Site Script ${props.siteScript.Id} could not be loaded`, error);
        }).then(() => {
            setLoading(false);
        });
    }, [props.siteScript]);

    const onTitleChanged = (ev: any, title: string) => {
        setEditingSiteScript({ ...editingSiteScript, Title: title });
    };

    const onDescriptionChanged = (ev: any, description: string) => {
        setEditingSiteScript({ ...editingSiteScript, Description: description });
    };

    const onVersionChanged = (ev: any, version: string) => {
        const versionInt = parseInt(version);
        if (!isNaN(versionInt)) {
            setEditingSiteScript({ ...editingSiteScript, Version: versionInt });
        }
    };

    const onSiteScriptContentUpdated = (siteScriptContent: ISiteScriptContent) => {
        setEditingSiteScript({ ...editingSiteScript, Content: siteScriptContent });
    };

    const onSave = async () => {
        setIsSaving(true);
        await siteDesignsService.saveSiteScript(editingSiteScript);
        const refreshedSiteScripts = await siteDesignsService.getSiteScripts();
        execute("SET_ALL_AVAILABLE_SITE_SCRIPTS", { siteScripts: refreshedSiteScripts } as ISetAllAvailableSiteScripts);
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
                    setEditingSiteScript({ ...editingSiteScript, Content: updatedScriptContent });
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
                            value={editingSiteScript.Title}
                            onChange={onTitleChanged} />
                    </div>
                    <div className={`${styles.column1} ${styles.righted}`}>
                        <PrimaryButton disabled={isSaving} text="Save" iconProps={{ iconName: "Save" }} onClick={() => onSave()} />
                    </div>
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
                            rows={1}
                            autoAdjustHeight
                            borderless
                            placeholder="Enter a description for the Site Script..."
                            onChange={onDescriptionChanged}
                        />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <Label>Actions</Label>
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.half}>
                        <SiteScriptDesigner
                            siteScriptContent={editingSiteScript.Content}
                            onSiteScriptContentUpdated={onSiteScriptContentUpdated} />
                    </div>
                    <div className={styles.half}>
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
                            value={siteScriptContentToCodeString(editingSiteScript.Content)}
                            editorDidMount={editorDidMount}
                        />
                    </div>
                </div>
            </div>
        </div>
    </div>;
};