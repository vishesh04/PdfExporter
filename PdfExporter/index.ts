import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import Auth from "./auth";
import FileUtils from "./fileUtils";

const auth = new Auth(process.env.GRAPH_CLIENT_ID, process.env.GRAPH_CLIENT_SECRET,
    process.env.GRAPH_TENANT_ID);

const fileUtils = new FileUtils(process.env.GRAPH_DRIVE_ID, process.env.GRAPH_WORKING_FOLDER_ID);

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    const fileUrl = req.query.source;
    const fileName = req.query.name;

    const accessToken = await auth.getAccessToken(true);
    const uploadSession = await fileUtils.getUploadSession(accessToken, fileName);
    const uploadUrl = uploadSession.data.uploadUrl;

    const downloadStream = (await fileUtils.downloadFile(fileUrl)).data;
    const fileSize = downloadStream.headers["content-length"];

    try {
        const response = await fileUtils.uploadFile(downloadStream, uploadUrl, fileSize);
        const pdfUrl = await fileUtils.getPdfUrl(response.id, accessToken);
        context.res = {
            body: { pdfUrl }
        };
    } catch (e) {
        console.log(e);
        context.res = {
            status: 500,
            body: {
                success: false
            }
        };
    }
    
};


export default httpTrigger;