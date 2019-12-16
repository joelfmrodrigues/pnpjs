import { SharePointQueryableCollection, SharePointQueryableInstance, defaultPath } from "./sharepointqueryable";
import { TextParser, BlobParser, JSONParser, BufferParser } from "@pnp/odata";
import { extend, getGUID, stringIsNullOrEmpty, jsS, isUrlAbsolute } from "@pnp/common";
import { LimitedWebPartManager } from "./webparts";
import { Item } from "./items";
import { SharePointQueryableShareableFile } from "./sharepointqueryableshareable";
import { odataUrlFrom } from "./odata";
import { extractWebUrl } from "./utils/extractweburl";

export interface ChunkedFileUploadProgressData {
    uploadId: string;
    stage: "starting" | "continue" | "finishing";
    blockNumber: number;
    totalBlocks: number;
    chunkSize: number;
    currentPointer: number;
    fileSize: number;
}

export interface AddUsingPathProps {
    /**
     * Overwrite the file if it exists
     */
    Overwrite: boolean;
    /**
     * specifies whether to auto checkout on invalid Data. It'll be useful if the list contains validation whose requirements upload will not be able to meet.
     */
    AutoCheckoutOnInvalidData?: boolean;
    /**
     * Specifies a XOR hash of the file data which should be used to ensure end-2-end data integrity, base64 representation
     */
    XorHash?: string;
}

/**
 * Describes a collection of File objects
 *
 */
@defaultPath("files")
export class Files extends SharePointQueryableCollection {

    /**
     * Gets a File by filename
     *
     * @param name The name of the file, including extension.
     */
    public getByName(name: string): File {
        const f = new File(this);
        f.concat(`('${name}')`);
        return f;
    }

    /**
     * Uploads a file. Not supported for batching
     *
     * @param url The folder-relative url of the file.
     * @param content The file contents blob.
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
     * @returns The new File and the raw response.
     */
    public add(url: string, content: string | ArrayBuffer | Blob, shouldOverWrite = true): Promise<FileAddResult> {
        return new Files(this, `add(overwrite=${shouldOverWrite},url='${url}')`)
            .postCore({
                body: content,
            }).then((response) => {
                return {
                    data: response,
                    file: this.getByName(url),
                };
            });
    }

    /**
     * Adds a file using the pound percent safe methods
     * 
     * @param url Excoded url of the file
     * @param content The file content
     * @param parameters Additional parameters to control method behavior
     */
    public addUsingPath(url: string, content: string | ArrayBuffer | Blob, parameters: AddUsingPathProps = { Overwrite: false }): Promise<FileAddResult> {

        const path = [`AddUsingPath(decodedurl='${url}'`];

        if (parameters) {
            if (parameters.Overwrite) {
                path.push(",Overwrite=true");
            }
            if (parameters.AutoCheckoutOnInvalidData) {
                path.push(",AutoCheckoutOnInvalidData=true");
            }
            if (!stringIsNullOrEmpty(parameters.XorHash)) {
                path.push(`,XorHash=${parameters.XorHash}`);
            }
        }

        path.push(")");

        return new Files(this, path.join(""))
            .postCore({
                body: content,
            }).then((response) => {
                return {
                    data: response,
                    file: this.getByName(url),
                };
            });
    }

    /**
     * Uploads a file. Not supported for batching
     *
     * @param url The folder-relative url of the file.
     * @param content The Blob file content to add
     * @param progress A callback function which can be used to track the progress of the upload
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
     * @param chunkSize The size of each file slice, in bytes (default: 10485760)
     * @returns The new File and the raw response.
     */
    public addChunked(
        url: string,
        content: Blob,
        progress?: (data: ChunkedFileUploadProgressData) => void,
        shouldOverWrite = true,
        chunkSize = 10485760,
    ): Promise<FileAddResult> {
        const adder = this.clone(Files, `add(overwrite = ${shouldOverWrite}, url = '${url}')`, false);
        return adder.postCore()
            .then(() => this.getByName(url))
            .then(file => file.setContentChunked(content, progress, chunkSize));
    }

    /**
     * Adds a ghosted file to an existing list or document library. Not supported for batching.
     *
     * @param fileUrl The server-relative url where you want to save the file.
     * @param templateFileType The type of use to create the file.
     * @returns The template file that was added and the raw response.
     */
    public addTemplateFile(fileUrl: string, templateFileType: TemplateFileType): Promise<FileAddResult> {
        return this.clone(Files, `addTemplateFile(urloffile = '${fileUrl}', templatefiletype = ${templateFileType})`, false)
            .postCore().then((response) => {
                return {
                    data: response,
                    file: this.getByName(fileUrl),
                };
            });
    }
}

/**
 * Describes a single File instance
 *
 */
export class File extends SharePointQueryableShareableFile {

    /**
     * Gets a value that specifies the list item field values for the list item corresponding to the file.
     *
     */
    public get listItemAllFields(): SharePointQueryableInstance {
        return new SharePointQueryableInstance(this, "listItemAllFields");
    }

    /**
     * Gets a collection of versions
     *
     */
    public get versions(): Versions {
        return new Versions(this);
    }

    /**
     * Approves the file submitted for content approval with the specified comment.
     * Only documents in lists that are enabled for content approval can be approved.
     *
     * @param comment The comment for the approval.
     */
    public approve(comment = ""): Promise<void> {
        return this.clone(File, `approve(comment = '${comment}')`).postCore();
    }

    /**
     * Stops the chunk upload session without saving the uploaded data. Does not support batching.
     * If the file doesn’t already exist in the library, the partially uploaded file will be deleted.
     * Use this in response to user action (as in a request to cancel an upload) or an error or exception.
     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
     * This method is currently available only on Office 365.
     *
     * @param uploadId The unique identifier of the upload session.
     */
    public cancelUpload(uploadId: string): Promise<void> {
        return this.clone(File, `cancelUpload(uploadId = guid'${uploadId}')`, false).postCore();
    }

    /**
     * Checks the file in to a document library based on the check-in type.
     *
     * @param comment A comment for the check-in. Its length must be <= 1023.
     * @param checkinType The check-in type for the file.
     */
    public checkin(comment = "", checkinType = CheckinType.Major): Promise<void> {

        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }

        return this.clone(File, `checkin(comment = '${comment}', checkintype = ${checkinType})`).postCore();
    }

    /**
     * Checks out the file from a document library.
     */
    public checkout(): Promise<void> {
        return this.clone(File, "checkout").postCore();
    }

    /**
     * Copies the file to the destination url.
     *
     * @param url The absolute url or server relative url of the destination file path to copy to.
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
     */
    public copyTo(url: string, shouldOverWrite = true): Promise<void> {
        return this.clone(File, `copyTo(strnewurl = '${url}', boverwrite = ${shouldOverWrite})`).postCore();
    }

    /**
     * Copies the file by path to destination path
     *
     * @param destUrl The absolute url or server relative url of the destination file path to copy to.
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
     * @param keepBoth Keep both if file with the same name in the same location already exists? Only relevant when shouldOverWrite is set to false.
     */
    public copyByPath(destUrl: string, shouldOverWrite: boolean, KeepBoth = false): Promise<void> {
        return this.select("ServerRelativeUrl").get().then(({ ServerRelativeUrl: srcUrl, ["odata.id"]: absoluteUrl }) => {
            const webBaseUrl = extractWebUrl(absoluteUrl);
            const hostUrl = webBaseUrl.replace("://", "___").split("/")[0].replace("___", "://");
            const f = new File(webBaseUrl, `/_api/SP.MoveCopyUtil.CopyFileByPath(overwrite=@a1)?@a1=${shouldOverWrite}`);
            return f.postCore({
                body: jsS({
                    destPath: {
                        DecodedUrl: isUrlAbsolute(destUrl) ? destUrl : `${hostUrl}${destUrl}`,
                        __metadata: {
                            type: "SP.ResourcePath",
                        },
                    },
                    options: {
                        KeepBoth: KeepBoth,
                        ResetAuthorAndCreatedOnCopy: true,
                        ShouldBypassSharedLocks: true,
                        __metadata: {
                            type: "SP.MoveCopyOptions",
                        },
                    },
                    srcPath: {
                        DecodedUrl: `${hostUrl}${srcUrl}`,
                        __metadata: {
                            type: "SP.ResourcePath",
                        },
                    },
                }),
            });
        });
    }

    /**
     * Delete this file.
     *
     * @param eTag Value used in the IF-Match header, by default "*"
     */
    public delete(eTag = "*"): Promise<void> {
        return this.clone(File, null).postCore({
            headers: {
                "IF-Match": eTag,
                "X-HTTP-Method": "DELETE",
            },
        });
    }

    /**
     * Denies approval for a file that was submitted for content approval.
     * Only documents in lists that are enabled for content approval can be denied.
     *
     * @param comment The comment for the denial.
     */
    public deny(comment = ""): Promise<void> {
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return this.clone(File, `deny(comment = '${comment}')`).postCore();
    }

    /**
     * Specifies the control set used to access, modify, or add Web Parts associated with this Web Part Page and view.
     * An exception is thrown if the file is not an ASPX page.
     *
     * @param scope The WebPartsPersonalizationScope view on the Web Parts page.
     */
    public getLimitedWebPartManager(scope = WebPartsPersonalizationScope.Shared): LimitedWebPartManager {
        return new LimitedWebPartManager(this, `getLimitedWebPartManager(scope = ${scope})`);
    }

    /**
     * Moves the file to the specified destination url.
     *
     * @param url The absolute url or server relative url of the destination file path to move to.
     * @param moveOperations The bitwise MoveOperations value for how to move the file.
     */
    public moveTo(url: string, moveOperations = MoveOperations.Overwrite): Promise<void> {
        return this.clone(File, `moveTo(newurl = '${url}', flags = ${moveOperations})`).postCore();
    }

    /**
     * Moves the file by path to the specified destination url.
     *
     * @param destUrl The absolute url or server relative url of the destination file path to move to.
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
     * @param keepBoth Keep both if file with the same name in the same location already exists? Only relevant when shouldOverWrite is set to false.
     */
    public moveByPath(destUrl: string, shouldOverWrite: boolean, KeepBoth = false): Promise<void> {
        return this.select("ServerRelativeUrl").get().then(({ ServerRelativeUrl: srcUrl, ["odata.id"]: absoluteUrl }) => {
            const webBaseUrl = extractWebUrl(absoluteUrl);
            const hostUrl = webBaseUrl.replace("://", "___").split("/")[0].replace("___", "://");
            const f = new File(webBaseUrl, `/_api/SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=${shouldOverWrite}`);
            return f.postCore({
                body: jsS({
                    destPath: {
                        DecodedUrl: isUrlAbsolute(destUrl) ? destUrl : `${hostUrl}${destUrl}`,
                        __metadata: {
                            type: "SP.ResourcePath",
                        },
                    },
                    options: {
                        KeepBoth: KeepBoth,
                        ResetAuthorAndCreatedOnCopy: false,
                        ShouldBypassSharedLocks: true,
                    },
                    srcPath: {
                        DecodedUrl: `${hostUrl}${srcUrl}`,
                        __metadata: {
                            type: "SP.ResourcePath",
                        },
                    },
                }),
            });
        });
    }

    /**
     * Submits the file for content approval with the specified comment.
     *
     * @param comment The comment for the published file. Its length must be <= 1023.
     */
    public publish(comment = ""): Promise<void> {
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return this.clone(File, `publish(comment = '${comment}')`).postCore();
    }

    /**
     * Moves the file to the Recycle Bin and returns the identifier of the new Recycle Bin item.
     *
     * @returns The GUID of the recycled file.
     */
    public recycle(): Promise<string> {
        return this.clone(File, "recycle").postCore();
    }

    /**
     * Reverts an existing checkout for the file.
     *
     */
    public undoCheckout(): Promise<void> {
        return this.clone(File, "undoCheckout").postCore();
    }

    /**
     * Removes the file from content approval or unpublish a major version.
     *
     * @param comment The comment for the unpublish operation. Its length must be <= 1023.
     */
    public unpublish(comment = ""): Promise<void> {
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return this.clone(File, `unpublish(comment = '${comment}')`).postCore();
    }

    /**
     * Gets the contents of the file as text. Not supported in batching.
     *
     */
    public getText(): Promise<string> {

        return this.clone(File, "$value", false).get(new TextParser(), { headers: { "binaryStringResponseBody": "true" } });
    }

    /**
     * Gets the contents of the file as a blob, does not work in Node.js. Not supported in batching.
     *
     */
    public getBlob(): Promise<Blob> {

        return this.clone(File, "$value", false).get(new BlobParser(), { headers: { "binaryStringResponseBody": "true" } });
    }

    /**
     * Gets the contents of a file as an ArrayBuffer, works in Node.js. Not supported in batching.
     */
    public getBuffer(): Promise<ArrayBuffer> {

        return this.clone(File, "$value", false).get(new BufferParser(), { headers: { "binaryStringResponseBody": "true" } });
    }

    /**
     * Gets the contents of a file as an ArrayBuffer, works in Node.js. Not supported in batching.
     */
    public getJSON(): Promise<any> {

        return this.clone(File, "$value", false).get(new JSONParser(), { headers: { "binaryStringResponseBody": "true" } });
    }

    /**
     * Sets the content of a file, for large files use setContentChunked. Not supported in batching.
     *
     * @param content The file content
     *
     */
    public setContent(content: string | ArrayBuffer | Blob): Promise<File> {

        return this.clone(File, "$value", false).postCore({
            body: content,
            headers: {
                "X-HTTP-Method": "PUT",
            },
        }).then(_ => new File(this));
    }

    /**
     * Gets the associated list item for this folder, loading the default properties
     */
    public getItem<T>(...selects: string[]): Promise<Item & T> {

        const q = this.listItemAllFields;
        return q.select.apply(q, selects).get().then((d: any) => {

            return extend((new Item(odataUrlFrom(d))).configureFrom(this), d);
        });
    }

    /**
     * Sets the contents of a file using a chunked upload approach. Not supported in batching.
     *
     * @param file The file to upload
     * @param progress A callback function which can be used to track the progress of the upload
     * @param chunkSize The size of each file slice, in bytes (default: 10485760)
     */
    public setContentChunked(file: Blob, progress?: (data: ChunkedFileUploadProgressData) => void, chunkSize = 10485760): Promise<FileAddResult> {

        if (progress === undefined) {
            progress = () => null;
        }

        const fileSize = file.size;
        const blockCount = parseInt((file.size / chunkSize).toString(), 10) + ((file.size % chunkSize === 0) ? 1 : 0);
        const uploadId = getGUID();

        // start the chain with the first fragment
        progress({ uploadId, blockNumber: 1, chunkSize, currentPointer: 0, fileSize, stage: "starting", totalBlocks: blockCount });

        let chain = this.startUpload(uploadId, file.slice(0, chunkSize));

        // skip the first and last blocks
        for (let i = 2; i < blockCount; i++) {
            chain = chain.then(pointer => {
                progress({ uploadId, blockNumber: i, chunkSize, currentPointer: pointer, fileSize, stage: "continue", totalBlocks: blockCount });
                return this.continueUpload(uploadId, pointer, file.slice(pointer, pointer + chunkSize));
            });
        }

        return chain.then(pointer => {
            progress({ uploadId, blockNumber: blockCount, chunkSize, currentPointer: pointer, fileSize, stage: "finishing", totalBlocks: blockCount });
            return this.finishUpload(uploadId, pointer, file.slice(pointer));
        });
    }

    /**
     * Starts a new chunk upload session and uploads the first fragment.
     * The current file content is not changed when this method completes.
     * The method is idempotent (and therefore does not change the result) as long as you use the same values for uploadId and stream.
     * The upload session ends either when you use the CancelUpload method or when you successfully
     * complete the upload session by passing the rest of the file contents through the ContinueUpload and FinishUpload methods.
     * The StartUpload and ContinueUpload methods return the size of the running total of uploaded data in bytes,
     * so you can pass those return values to subsequent uses of ContinueUpload and FinishUpload.
     * This method is currently available only on Office 365.
     *
     * @param uploadId The unique identifier of the upload session.
     * @param fragment The file contents.
     * @returns The size of the total uploaded data in bytes.
     */
    protected startUpload(uploadId: string, fragment: ArrayBuffer | Blob): Promise<number> {
        return this.clone(File, `startUpload(uploadId = guid'${uploadId}')`, false)
            .postCore<string>({ body: fragment })
            .then(n => {
                // When OData=verbose the payload has the following shape:
                // { StartUpload: "10485760" }
                if (typeof n === "object") {
                    n = (n as any).StartUpload;
                }
                return parseFloat(n);
            });
    }

    /**
     * Continues the chunk upload session with an additional fragment.
     * The current file content is not changed.
     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
     * This method is currently available only on Office 365.
     *
     * @param uploadId The unique identifier of the upload session.
     * @param fileOffset The size of the offset into the file where the fragment starts.
     * @param fragment The file contents.
     * @returns The size of the total uploaded data in bytes.
     */
    protected continueUpload(uploadId: string, fileOffset: number, fragment: ArrayBuffer | Blob): Promise<number> {
        return this.clone(File, `continueUpload(uploadId = guid'${uploadId}', fileOffset = ${fileOffset})`, false)
            .postCore<string>({ body: fragment })
            .then(n => {
                // When OData=verbose the payload has the following shape:
                // { ContinueUpload: "20971520" }
                if (typeof n === "object") {
                    n = (n as any).ContinueUpload;
                }
                return parseFloat(n);
            });
    }

    /**
     * Uploads the last file fragment and commits the file. The current file content is changed when this method completes.
     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
     * This method is currently available only on Office 365.
     *
     * @param uploadId The unique identifier of the upload session.
     * @param fileOffset The size of the offset into the file where the fragment starts.
     * @param fragment The file contents.
     * @returns The newly uploaded file.
     */
    protected finishUpload(uploadId: string, fileOffset: number, fragment: ArrayBuffer | Blob): Promise<FileAddResult> {
        return this.clone(File, `finishUpload(uploadId = guid'${uploadId}', fileOffset = ${fileOffset})`, false)
            .postCore<{ ServerRelativeUrl: string }>({ body: fragment })
            .then(response => {
                return {
                    data: response,
                    file: new File(odataUrlFrom(response)),
                };
            });
    }
}

/**
 * Describes a collection of Version objects
 *
 */
@defaultPath("versions")
export class Versions extends SharePointQueryableCollection {

    /**	
     * Gets a version by id	
     *	
     * @param versionId The id of the version to retrieve	
     */
    public getById(versionId: number): Version {
        const v = new Version(this);
        v.concat(`(${versionId})`);
        return v;
    }

    /**
     * Deletes all the file version objects in the collection.
     *
     */
    public deleteAll(): Promise<void> {
        return new Versions(this, "deleteAll").postCore();
    }

    /**
     * Deletes the specified version of the file.
     *
     * @param versionId The ID of the file version to delete.
     */
    public deleteById(versionId: number): Promise<void> {
        return this.clone(Versions, `deleteById(vid = ${versionId})`).postCore();
    }

    /**
     * Recycles the specified version of the file.
     *
     * @param versionId The ID of the file version to delete.
     */
    public recycleByID(versionId: number): Promise<void> {
        return this.clone(Versions, `recycleByID(vid = ${versionId})`).postCore();
    }

    /**
     * Deletes the file version object with the specified version label.
     *
     * @param label The version label of the file version to delete, for example: 1.2
     */
    public deleteByLabel(label: string): Promise<void> {
        return this.clone(Versions, `deleteByLabel(versionlabel = '${label}')`).postCore();
    }

    /**
     * Recycles the file version object with the specified version label.
     *
     * @param label The version label of the file version to delete, for example: 1.2
     */
    public recycleByLabel(label: string): Promise<void> {
        return this.clone(Versions, `recycleByLabel(versionlabel = '${label}')`).postCore();
    }

    /**
     * Creates a new file version from the file specified by the version label.
     *
     * @param label The version label of the file version to restore, for example: 1.2
     */
    public restoreByLabel(label: string): Promise<void> {
        return this.clone(Versions, `restoreByLabel(versionlabel = '${label}')`).postCore();
    }
}

/**
 * Describes a single Version instance
 *
 */
export class Version extends SharePointQueryableInstance {

    /**
    * Delete a specific version of a file.
    *
    * @param eTag Value used in the IF-Match header, by default "*"
    */
    public delete = this._deleteWithETag;

    // /**
    // * Delete a specific version of a file.
    // *
    // * @param eTag Value used in the IF-Match header, by default "*"
    // */
    // public delete(eTag = "*"): Promise<void> {
    //     return this.postCore({
    //         headers: {
    //             "IF-Match": eTag,
    //             "X-HTTP-Method": "DELETE",
    //         },
    //     });
    // }
}

export enum CheckinType {
    Minor = 0,
    Major = 1,
    Overwrite = 2,
}

export interface FileAddResult {
    file: File;
    data: any;
}

export enum WebPartsPersonalizationScope {
    User = 0,
    Shared = 1,
}

export enum MoveOperations {
    Overwrite = 1,
    AllowBrokenThickets = 8,
}

export enum TemplateFileType {
    StandardPage = 0,
    WikiPage = 1,
    FormPage = 2,
    ClientSidePage = 3,
}
