const express = require("express");
const { google } = require("googleapis");
const app = express();
app.set("view engine", "ejs");
app.use(express.urlencoded({ extended: true }));

function manychatjson(json,cardnum=10) {
  if (cardnum>num || cardnum == 0)
      cardnum = num;
  //{if you want to edit} edit the num varibles to the length of the list you want to repart about
  var num = json["values"].length,i;
  manychat = {"version": "v2","content": {"messages": [],"actions": [],"quick_replies": []}};
  for (i = 0; i < Math.floor(num/cardnum); i++)   {
      manychat["content"]["messages"].push({"type": "cards","elements": [],"image_aspect_ratio": "horizontal"});
      for (let j = 0; j < cardnum; j++)   {
          /*edit each vaible the way you want it to display
          don't forget the other for loop*/
          title = json["value"][(i*cardnum)+j][ProductName];
          subtitle = json["value"][(i*cardnum)+j][ProductDiscreption];
          image_url = json["value"][(i*cardnum)+j][ProductImageURL];
          manychat["content"]["messages"][i]["elements"].push({"title": title,"subtitle": subtitle,"image_url": image_url,"action_url": "https://manychat.com","buttons": []});
      }
  }
  if (num%cardnum > 0)   {
      manychat["content"]["messages"].push({"type": "cards","elements": [],"image_aspect_ratio": "horizontal"});
          for(let j = 0; j < num%cardnum; j++)    {
              //this one
              title = json["value"][(i*cardnum)+j][ProductName];
              subtitle = json["value"][(i*cardnum)+j][ProductDiscreption];
              image_url = json["value"][(i*cardnum)+j][ProductImageURL];
              manychat["content"]["messages"][i]["elements"].push({"title": title,"subtitle": subtitle,"image_url": image_url,"action_url": "https://manychat.com","buttons": []});    
          }
  }
  return manychat;
}

app.get("/get", async (req, res) => {
  const { request, name } = req.body;
  const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });
  // Create client instance for auth
  const client = await auth.getClient();
  // Instance of Google Sheets API
  const googleSheets = google.sheets({ version: "v4", auth: client });
  const spreadsheetId = "1L5VOht4Rw7tYeNh6Dmgtg25acYN3TB5WEDAUhJz7Bd4";
  const getRows = await googleSheets.spreadsheets.values.get({
    auth,
    spreadsheetId,
    range: "Products!A2:Z1000",
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  x = getRows.data.values
  console.log(x.length)
  var xx = {"value" : []};
  for (let i = 0; i < x.length; i++) {
    xx["value"].push({
      "ProductID" : x[i][0],
      "Active": x[i][1],
      "ProductName" : x[i][2],
      "Cats" : x[i][3],
      "ProductDiscreption" : x[i][4],
      "ProductPrice" : x[i][5],
      "Pricing" : x[i][6],
      "Time" : x[i][7],
      "ProductImageURL" : x[i][8]})
  }
  var filter = xx.value.filter(function (el) {
    return el.Cats == 'لافيتات';
  });
  res.send(filter);
});

app.get("/append", async (req, res) => {
  const { request, name } = req.body;

  const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });

  // Create client instance for auth
  const client = await auth.getClient();

  // Instance of Google Sheets API
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


app.listen(1337, (req, res) => console.log("running on 1337"));