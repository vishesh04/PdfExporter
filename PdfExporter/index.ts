import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import Auth from "./auth";
import FileUtils from "./fileUtils";

const auth = new Auth(process.env.GRAPH_CLIENT_ID, process.env.GRAPH_CLIENT_SECRET,
    process.env.GRAPH_TENANT_ID);

const fileUtils = new FileUtils(process.env.GRAPH_DRIVE_ID, process.env.GRAPH_WORKING_FOLDER_ID);

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    const accessToken = await auth.getAccessToken(true);

    if (req.method === "DELETE") {
        const fileId = context.bindingData.id;
        await fileUtils.deleteFile(fileId, accessToken);
        context.res = {
            body: {
                success: true
            }
        };
    } else {
        const fileUrl = req.query.source;
        const fileName = req.query.name;

        const uploadSession = await fileUtils.getUploadSession(accessToken, fileName);
        const uploadUrl = uploadSession.data.uploadUrl;
    
        const downloadStream = (await fileUtils.downloadFile(fileUrl)).data;
        const fileSize = downloadStream.headers["content-length"];
    
        try {
            const response = await fileUtils.uploadFile(downloadStream, uploadUrl, fileSize);
            const pdfUrl = await fileUtils.getPdfUrl(response.id, accessToken);
            context.res = {
                body: { 
                    pdfUrl,
                    tempFileId: response.id
                }
            };
        } catch (e) {
            context.log(e);
            context.res = {
                status: 500,
                body: {
                    success: false
                }
            };
        }
    }
};


export default httpTrigger;