import express = require('express');
const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
let PORT = process.env.PORT || 5000;
import * as requestp from "request-promise";
import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';

type ManychatGallery = {
  version: string;
  content: {
    messages: Array<{
      type: string;
      elements: Array<{
        title: string;
        subtitle: string;
        image_url: string;
        action_url: string;
        buttons: Array<any>;
      }>;
      image_aspect_ratio?: string;
    }>;
    actions: Array<any>;
    quick_replies: Array<any>;
  };
}
type SharePointListItems = {
  value: Array<{
    FileSystemObjectType?: number
    Id?: number
    ServerRedirectedEmbedUri: any
    ServerRedirectedEmbedUrl?: string
    ID?: number
    ContentTypeId?: string
    Title?: string
    Modified?: string
    Created?: string
    AuthorId?: number
    EditorId?: number
    OData__UIVersionString?: string
    Attachments?: boolean
    GUID?: string
    ComplianceAssetId: any
    Active?: boolean
    ProductName: string
    Cats: string
    ProductDiscreption: string
    ProductPrice: number
    Pricing?: string
    Time?: string
    ProductImageURL: string
  }>
}
type GoogleSheet = {
  value: Array<{
    ProductID: string;
    Active: string;
    ProductName: string;
    Cats: string;
    ProductDiscreption: string;
    ProductPrice: string;
    Pricing: string;
    Time?: string; // optional property
    ProductImageURL: string;
    __PowerAppsId__: string;
  }>;
}



function DarAssendangoogle(json: GoogleSheet, cardnum = 10) {
  let num: number = json.value.length, i;
  let title: string, subtitle: string, image_url: string;
  let manychat:ManychatGallery = { "version": "v2", "content": { "messages": [], "actions": [], "quick_replies": [] } };
  for (i = 0; i < Math.floor(num / cardnum); i++) {
    manychat["content"]["messages"].push({ "type": "cards", "elements": [], "image_aspect_ratio": "horizontal" });
    for (let j = 0; j < cardnum; j++) {
      /*edit each vaible the way you want it to display
      don't forget the other for loop*/
      title = json.value[(i * cardnum) + j]["ProductName"];
      subtitle = json.value[(i * cardnum) + j]["ProductDiscreption"];
      image_url = json.value[(i * cardnum) + j]["ProductImageURL"];
      manychat["content"]["messages"][i]["elements"].push({ "title": title, "subtitle": subtitle, "image_url": image_url, "action_url": "https://manychat.com", "buttons": [] });
    }
  }
  if (num % cardnum > 0) {
    manychat["content"]["messages"].push({ "type": "cards", "elements": [], "image_aspect_ratio": "horizontal" });
    for (let j = 0; j < num % cardnum; j++) {
      //this one
      title = json.value[(i * cardnum) + j]["ProductName"];
      subtitle = json.value[(i * cardnum) + j]["ProductDiscreption"];
      image_url = json.value[(i * cardnum) + j]["ProductImageURL"];
      manychat["content"]["messages"][i]["elements"].push({ "title": title, "subtitle": subtitle, "image_url": image_url, "action_url": "https://manychat.com", "buttons": [] });
    }
  }
  return manychat;
}

app.get("/DarAssendan/Google/get", async (req, res) => {
  const auth: GoogleAuth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });
  const client = await auth.getClient();

  const googleSheets = google.sheets({ version: "v4", auth: client });
  const spreadsheetId: string = "1L5VOht4Rw7tYeNh6Dmgtg25acYN3TB5WEDAUhJz7Bd4";
  const getRows = await googleSheets.spreadsheets.values.get({
    auth,
    spreadsheetId,
    range: "Products!A1:Z1000",
  });
  let googledata: string[][];
  if (getRows.data.values == undefined) {
    res.send({});
    return;
  }
  else {
    googledata = getRows.data.values;
  }
  let map: GoogleSheet = { "value": [] };
  for (let i = 1; i < googledata.length; i++) {
    let temp: any = {};
    for (let j = 0; j < googledata[0].length; j++) {
      temp[googledata[0][j]] = googledata[i][j];
    }
    map["value"].push(temp)
  }
  // var manychatfilter = map.value.filter(function (el) {
  //  return el.Cats == req.body.Categories;
  //});
  res.send(DarAssendangoogle(map));
});

// use the custom interfaces in your function signature
function DarAssendansharepoint(json: SharePointListItems, cardnum = 10): ManychatGallery {
  let num: number = json["value"].length, i: number;
  let title: string, subtitle: string, image_url: string;
  const manychat: ManychatGallery = { "version": "v2", "content": { "messages": [], "actions": [], "quick_replies": [] } };
  for (i = 0; i < Math.floor(num / cardnum); i++) {
    manychat["content"]["messages"].push({ "type": "cards", "elements": [], "image_aspect_ratio": "horizontal" });
    for (let j = 0; j < cardnum; j++) {
      /*edit each vaible the way you want it to display
      don't forget the other for loop*/
      title = json["value"][(i * cardnum) + j]["ProductName"];
      subtitle = json["value"][(i * cardnum) + j]["ProductDiscreption"];
      image_url = json["value"][(i * cardnum) + j]["ProductImageURL"];
      manychat["content"]["messages"][i]["elements"].push({ "title": title, "subtitle": subtitle, "image_url": image_url, "action_url": "https://manychat.com", "buttons": [] });
    }
  }
  if (num % cardnum > 0) {
    manychat["content"]["messages"].push({ "type": "cards", "elements": [], "image_aspect_ratio": "horizontal" });
    for (let j = 0; j < num % cardnum; j++) {
      //this one
      title = json["value"][(i * cardnum) + j]["ProductName"];
      subtitle = json["value"][(i * cardnum) + j]["ProductDiscreption"];
      image_url = json["value"][(i * cardnum) + j]["ProductImageURL"];
      manychat["content"]["messages"][i]["elements"].push({ "title": title, "subtitle": subtitle, "image_url": image_url, "action_url": "https://manychat.com", "buttons": [] });
    }
  }
  return manychat;
}
class RedTech {
  static TenantID: string = "e1cbaea0-f8c8-4ce7-a484-ade672f55f8d";
  static TenantName: string = "redtechly0";
  static ApplicationID: string = "00000003-0000-0ff1-ce00-000000000000";
  static ClientID: string = "bc1fdfac-50ef-4031-90c8-f12401b38e7a";
  static ClientSecret: string = "br+8R9XC+97p24aP8iSZaIgl16CKOxUjpcVlG41WXnM=";
  static RefreshToken: string = "PAQABAAEAAAD--DLA3VO7QrddgJg7WevrLVZx_-lO6VBwT_3LyQbCKCz_y31k2_NBUlGOCHVaTQNssNFOz9ciOG68LVin33nflACoDGbMXjK6GMvshqE6z7ccWyDM2Rm4Nc9WJUJazyiIxg1xCFWRNm2kwLtsw4klNSe9kEtoJQyAMTxRfqpTBn_RDXQYM7NCGkczXQ_r4V23GesvEFuggk5JDi3r9N-6Rr4C7VU7JQaX5VWn4tWDmdiGQC8d4b-lIbBMN6Iwm8dCmjQw7tenNj-K7sLKt4QioI31zIt46Q6MxsRkjUFCqCAA";
  static AccessToken: string = "";
  static IssueDate: number = 0;
  SiteName: string;
  ListName: string;
  constructor(ListName: string, SiteName: string) {
    this.SiteName = SiteName
    this.ListName = ListName
  }
  static async generatetoken() {
    let result: string = await requestp.post({
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      form: {
        grant_type: "refresh_token",
        client_id: this.ClientID + "@" + this.TenantID,
        client_secret: this.ClientSecret,
        resource: this.ApplicationID + "/" + this.TenantName + ".sharepoint.com@" + this.TenantID,
        refresh_token: this.RefreshToken
      },
      uri: "https://accounts.accesscontrol.windows.net/" + this.TenantID + "/tokens/OAuth/2"
    });
    return JSON.parse(result).access_token;
  }
  static async getAccessToken() {
    if (this.AccessToken == null || Date.now().valueOf() > (this.IssueDate + 28500000)) {
      console.log("Generating new token " + ++NumOfToken);
      this.IssueDate = Date.now().valueOf();
      this.AccessToken = await this.generatetoken();
    }
    return this.AccessToken;
  }
  async getListItems() {
    let result = await requestp.get({
      json: true,
      headers: {
        "Authorization": "Bearer " + await RedTech.getAccessToken(),
        "Content-Type": "application/json; odata=verbose",
        "Accept": "application/json; odata=nometadata"
      },
      uri: encodeURI("https://" + RedTech.TenantName + ".sharepoint.com/sites/" + this.SiteName + "/_api/web/lists/GetByTitle('" + this.ListName + "')/items"),
      body: ""
    });
    return result;
  }

}

app.get("/DarAssendan/SharePoint/get", async (req:express.Request, res:express.Response) => {
  let SiteName: string = "REDCompanies";
  let ListName: string = "DarAssendan";
  let darassedn: RedTech = new RedTech(ListName, SiteName)
  res.send(DarAssendansharepoint(await darassedn.getListItems()));
});

app.listen(PORT, () => {
  console.log(`listening on post http://localhost:${PORT}`);
});

