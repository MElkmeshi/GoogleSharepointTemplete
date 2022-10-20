const express = require("express");
const { google } = require("googleapis");
const app = express();
let PORT = process.env.PORT || 5000;
app.set("view engine", "ejs");
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
var request = require('request');
var FormData = require('form-data');
var fs = require('fs');

//DarAssendan

//first googlesheets
function DarAssendangoogle(json,cardnum=10) {
  if (cardnum>num || cardnum == 0)
      cardnum = num;
  //{if you want to edit} edit the num varibles to the length of the list you want to repart about
  var num = json.length,i;
  manychat = {"version": "v2","content": {"messages": [],"actions": [],"quick_replies": []}};
  for (i = 0; i < Math.floor(num/cardnum); i++)   {
      manychat["content"]["messages"].push({"type": "cards","elements": [],"image_aspect_ratio": "horizontal"});
      for (let j = 0; j < cardnum; j++)   {
          /*edit each vaible the way you want it to display
          don't forget the other for loop*/
          title = json[(i*cardnum)+j]["ProductName"];
          subtitle = json[(i*cardnum)+j]["ProductDiscreption"];
          image_url = json[(i*cardnum)+j]["ProductImageURL"];
          manychat["content"]["messages"][i]["elements"].push({"title": title,"subtitle": subtitle,"image_url": image_url,"action_url": "https://manychat.com","buttons": []});
      }
  }
  if (num%cardnum > 0)   {
      manychat["content"]["messages"].push({"type": "cards","elements": [],"image_aspect_ratio": "horizontal"});
          for(let j = 0; j < num%cardnum; j++)    {
              //this one
              title = json[(i*cardnum)+j]["ProductName"];
              subtitle = json[(i*cardnum)+j]["ProductDiscreption"];
              image_url = json[(i*cardnum)+j]["ProductImageURL"];
              manychat["content"]["messages"][i]["elements"].push({"title": title,"subtitle": subtitle,"image_url": image_url,"action_url": "https://manychat.com","buttons": []});    
          }
  }
  return manychat;
}
app.post("/DarAssendan/Google/get", async (req, res) => {
  //filter based on Categories
  const { request, name } = req.body;
  const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });
  const client = await auth.getClient();
  const googleSheets = google.sheets({ version: "v4", auth: client });
  const spreadsheetId = "1L5VOht4Rw7tYeNh6Dmgtg25acYN3TB5WEDAUhJz7Bd4";
  const getRows = await googleSheets.spreadsheets.values.get({
    auth,
    spreadsheetId,
    range: "Products!A2:Z1000",
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  googledata = getRows.data.values
  var manychat = {"value" : []};
  for (let i = 0; i < googledata.length; i++) {
    manychat["value"].push({
      "ProductID" : googledata[i][0],
      "Active": googledata[i][1],
      "ProductName" : googledata[i][2],
      "Cats" : googledata[i][3],
      "ProductDiscreption" : googledata[i][4],
      "ProductPrice" : googledata[i][5],
      "Pricing" : googledata[i][6],
      "Time" : googledata[i][7],
      "ProductImageURL" : googledata[i][8]})
  }
  var manychatfilter = manychat.value.filter(function (el) {
    return el.Cats == req.body.Categories;
  });
  res.send(DarAssendangoogle(manychatfilter));
});
app.post("/DarAssendan/Google/append", async (req, res) => {
  const { request, name } = req.body;
  const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });
  const client = await auth.getClient();
  const googleSheets = google.sheets({ version: "v4", auth: client });
  const spreadsheetId = "1L5VOht4Rw7tYeNh6Dmgtg25acYN3TB5WEDAUhJz7Bd4";
  await googleSheets.spreadsheets.values.append({
    auth,
    spreadsheetId,
    range: "Questions!A:E",
    valueInputOption: "USER_ENTERED",
    resource: {
      values: [['2022-10-03T16:06:07.1015285Z','Messenger','','نبي نسال لو فيه طباعه درفت او زي كتالوج ديجيتل 50 قطعه','Hala Elamari' ]],
    },
  });
  res.send('Done');
});
//second Sharepoint
function manychatjson(json,cardnum=10) {
   
  if (cardnum>num || cardnum == 0)
      cardnum = num;
  //{if you want to edit} edit the num varibles to the length of the list you want to repart about
  var num = json["value"].length,i;
  manychat = {"version": "v2","content": {"messages": [],"actions": [],"quick_replies": []}};
  for (i = 0; i < Math.floor(num/cardnum); i++)   {
      manychat["content"]["messages"].push({"type": "cards","elements": [],"image_aspect_ratio": "horizontal"});
      for (let j = 0; j < cardnum; j++)   {
          /*edit each vaible the way you want it to display
          don't forget the other for loop*/
          title = json["value"][(i*cardnum)+j]["ProductName"];
          subtitle = json["value"][(i*cardnum)+j]["ProductDiscreption"];
          image_url = json["value"][(i*cardnum)+j]["ProductImageURL"];
          manychat["content"]["messages"][i]["elements"].push({"title": title,"subtitle": subtitle,"image_url": image_url,"action_url": "https://manychat.com","buttons": []});
      }
  }
  if (num%cardnum > 0)   {
      manychat["content"]["messages"].push({"type": "cards","elements": [],"image_aspect_ratio": "horizontal"});
          for(let j = 0; j < num%cardnum; j++)    {
              //this one
              title = json["value"][(i*cardnum)+j]["ProductName"];
              subtitle = json["value"][(i*cardnum)+j]["ProductDiscreption"];
              image_url = json["value"][(i*cardnum)+j]["ProductImageURL"]; 
              manychat["content"]["messages"][i]["elements"].push({"title": title,"subtitle": subtitle,"image_url": image_url,"action_url": "https://manychat.com","buttons": []});    
          }
  }
  return manychat;
}
function sh(callback){
  var x;
  ApplicationID = "00000003-0000-0ff1-ce00-000000000000";
  TenantName = "redtechly0";
  TenantID = "e1cbaea0-f8c8-4ce7-a484-ade672f55f8d";
  ClientID = "bc1fdfac-50ef-4031-90c8-f12401b38e7a";
  ClientSecret = "br+8R9XC+97p24aP8iSZaIgl16CKOxUjpcVlG41WXnM=";
  RefreshToken = "PAQABAAEAAAD--DLA3VO7QrddgJg7WevrLVZx_-lO6VBwT_3LyQbCKCz_y31k2_NBUlGOCHVaTQNssNFOz9ciOG68LVin33nflACoDGbMXjK6GMvshqE6z7ccWyDM2Rm4Nc9WJUJazyiIxg1xCFWRNm2kwLtsw4klNSe9kEtoJQyAMTxRfqpTBn_RDXQYM7NCGkczXQ_r4V23GesvEFuggk5JDi3r9N-6Rr4C7VU7JQaX5VWn4tWDmdiGQC8d4b-lIbBMN6Iwm8dCmjQw7tenNj-K7sLKt4QioI31zIt46Q6MxsRkjUFCqCAA";
  var x;
  request({
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      form: {
          grant_type: "refresh_token",
          client_id: ClientID + "@" + TenantID,
          client_secret: ClientSecret,  
          resource: ApplicationID + "/" + TenantName + ".sharepoint.com@" + TenantID,
          refresh_token: RefreshToken
      },
      uri: "https://accounts.accesscontrol.windows.net/" + TenantID + "/tokens/OAuth/2",
      method: 'POST'
  }, function (_err, _resreq, _body) {
     x = JSON.parse(_body);
     callback(x);
  }); 
}
app.post("/DarAssendan/SharePoint/get", (req,res) => {
  sh(function (body){
    var xzs = 'hi';
      request({
          json:true,
          headers: {
              "Authorization" : "Bearer "+ body.access_token,
              "Content-Type" : "application/json; odata=verbose",
              "Accept" : "application/json; odata=nometadata"
          },
          uri: encodeURI("https://redtechly0.sharepoint.com/sites/REDCompanies/_api/web/lists/GetByTitle('DarAssendan')/items?$filter=Cats eq " + "'" + req.body.Categories + "'"),
          body: "",
          method: 'GET'
        }, function (_err, _resreq, _body) {
          res.send(manychatjson(_body))
        })        
  });
});
app.post("/DarAssendan/SharePoint/append", (req,res) => {
  sh(function (body){
    var xzs = 'hi';
      request({
          json:true,
          headers: {
              "Authorization" : "Bearer "+ body.access_token,
              "Content-Type" : "application/json; odata=verbose",
              "Accept" : "application/json; odata=nometadata"
          },
          uri: encodeURI("https://redtechly0.sharepoint.com/sites/REDCompanies/_api/web/lists/GetByTitle('DarAssendan')/items?$filter=Cats eq " + req.body.Categories),
          body: "",
          method: 'GET'
        }, function (_err, _resreq, _body) {
          res.send(_body)
        })        
  });
});


app.get("/", async (req, res) => {
  res.send('Welcome to Red Tech MicroService');
});

app.listen(PORT, () => {
    console.log(`listening on post ${PORT}`);
});

