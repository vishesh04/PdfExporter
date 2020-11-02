import axios, { AxiosRequestConfig, AxiosPromise } from "axios";
import { resolve } from "path";

const API_BASE_URL = "https://graph.microsoft.com/v1.0";

export default class FileUtils {
  driveId: string;
  workingFolderId: string;

  constructor(driveId: string,workingFolderId: string ) {
    this.driveId = driveId;
    this.workingFolderId = workingFolderId;
  }

  async getUploadSession(accessToken: string, fileName: string) {
    return axios.post(`${API_BASE_URL}/drives/${this.driveId}/items/${this.workingFolderId}:/${fileName}:/createUploadSession`, {
      item: {
        "@microsoft.graph.conflictBehavior": "rename"
      }
    }, {
      headers: { "Authorization": `Bearer ${accessToken}` }
    });
  }

  async uploadChunk(uploadUrl: string, bytes: any, contentRange: string) {
    return axios.put(uploadUrl, bytes, {
      headers: {
        "Content-Length": bytes.length,
        "Content-Range": contentRange
      }
    });
  }

  async uploadFile(inputStream, uploadUrl: string, fileSize: number): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      let uploadedBytes = 0;
      let stagedChunk = Buffer.alloc(0);
  
      const idealChunksize = 3276800 // 10 chunks of 320 KiB each, ~3.2768 MB
      
      inputStream.on("data", async (chunk) => {
          stagedChunk = Buffer.concat([stagedChunk, chunk]);
          
          if (stagedChunk.length >= idealChunksize) {
              const chunkToUpload = stagedChunk.slice(0, idealChunksize);
              stagedChunk = stagedChunk.slice(idealChunksize);
              try {
                inputStream.pause();
                  const contentRange = `bytes ${uploadedBytes}-${uploadedBytes + chunkToUpload.length - 1}/${fileSize}`;
                  await this.uploadChunk(uploadUrl, chunkToUpload, contentRange);
                  uploadedBytes += chunkToUpload.length;
                  inputStream.resume();
              } catch (e) {
                inputStream.destroy();
                reject(new Error("Failed"))
              }
          }
      });
  
      inputStream.on("end", async (chunk) => {
           // upload the last chunk
          try {
              const contentRange = `bytes ${uploadedBytes}-${uploadedBytes + stagedChunk.length - 1}/${fileSize}`;
              const response = await this.uploadChunk(uploadUrl, stagedChunk, contentRange);
              // Convert the file to pdf
              resolve(response.data)
          } catch (e) {
            inputStream.destroy();
            reject(new Error("Failed"))
          }
      });
    });
  }

  async getPdfUrl(fileId: string, accessToken: string) {
    try {
      const response = await axios.get(`${API_BASE_URL}/drives/${this.driveId}/items/${fileId}/content?format=pdf`, {
        headers: { "Authorization": `Bearer ${accessToken}` },
        maxRedirects: 0
      });
    } catch(e) {
      if (e.response.status === 302) {
        return e.response.headers.location
      }
      throw e;
    }
  }

  async downloadFile(fileUrl: string) {
    return axios({
      method: 'get',
      url: fileUrl,
      responseType: 'stream'
    });
  }

}