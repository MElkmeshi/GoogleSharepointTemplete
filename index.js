"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var express = require("express");
var app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
var PORT = process.env.PORT || 5000;
var requestp = require("request-promise");
var googleapis_1 = require("googleapis");
function DarAssendangoogle(json, cardnum) {
    if (cardnum === void 0) { cardnum = 10; }
    // if (cardnum > num || cardnum == 0)
    // cardnum = num;
    //{if you want to edit} edit the num varibles to the length of the list you want to repart about
    var num = json.value.length, i;
    var title, subtitle, image_url;
    var manychat = { "version": "v2", "content": { "messages": [], "actions": [], "quick_replies": [] } };
    for (i = 0; i < Math.floor(num / cardnum); i++) {
        manychat["content"]["messages"].push({ "type": "cards", "elements": [], "image_aspect_ratio": "horizontal" });
        for (var j = 0; j < cardnum; j++) {
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
        for (var j = 0; j < num % cardnum; j++) {
            //this one
            title = json.value[(i * cardnum) + j]["ProductName"];
            subtitle = json.value[(i * cardnum) + j]["ProductDiscreption"];
            image_url = json.value[(i * cardnum) + j]["ProductImageURL"];
            manychat["content"]["messages"][i]["elements"].push({ "title": title, "subtitle": subtitle, "image_url": image_url, "action_url": "https://manychat.com", "buttons": [] });
        }
    }
    return manychat;
}
app.get("/DarAssendan/Google/get", function (req, res) { return __awaiter(void 0, void 0, void 0, function () {
    var auth, client, googleSheets, spreadsheetId, getRows, googledata, map, i, temp, j;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                auth = new googleapis_1.google.auth.GoogleAuth({
                    keyFile: "credentials.json",
                    scopes: "https://www.googleapis.com/auth/spreadsheets",
                });
                return [4 /*yield*/, auth.getClient()];
            case 1:
                client = _a.sent();
                googleSheets = googleapis_1.google.sheets({ version: "v4", auth: client });
                spreadsheetId = "1L5VOht4Rw7tYeNh6Dmgtg25acYN3TB5WEDAUhJz7Bd4";
                return [4 /*yield*/, googleSheets.spreadsheets.values.get({
                        auth: auth,
                        spreadsheetId: spreadsheetId,
                        range: "Products!A1:Z1000",
                    })];
            case 2:
                getRows = _a.sent();
                if (getRows.data.values == undefined) {
                    res.send({});
                    return [2 /*return*/];
                }
                else {
                    googledata = getRows.data.values;
                }
                map = { "value": [] };
                for (i = 1; i < googledata.length; i++) {
                    temp = {};
                    for (j = 0; j < googledata[0].length; j++) {
                        temp[googledata[0][j]] = googledata[i][j];
                    }
                    map["value"].push(temp);
                }
                // var manychatfilter = map.value.filter(function (el) {
                //  return el.Cats == req.body.Categories;
                //});
                res.send(DarAssendangoogle(map));
                return [2 /*return*/];
        }
    });
}); });
app.get("/", function (request, response) { return __awaiter(void 0, void 0, void 0, function () {
    var x;
    return __generator(this, function (_a) {
        x = requestp("https://jsonplaceholder.typicode.com/todos/1");
        response.send(x);
        return [2 /*return*/];
    });
}); });
// use the custom interfaces in your function signature
function DarAssendansharepoint(json, cardnum) {
    if (cardnum === void 0) { cardnum = 10; }
    var num = json["value"].length, i;
    var title, subtitle, image_url;
    var manychat = { "version": "v2", "content": { "messages": [], "actions": [], "quick_replies": [] } };
    for (i = 0; i < Math.floor(num / cardnum); i++) {
        manychat["content"]["messages"].push({ "type": "cards", "elements": [], "image_aspect_ratio": "horizontal" });
        for (var j = 0; j < cardnum; j++) {
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
        for (var j = 0; j < num % cardnum; j++) {
            //this one
            title = json["value"][(i * cardnum) + j]["ProductName"];
            subtitle = json["value"][(i * cardnum) + j]["ProductDiscreption"];
            image_url = json["value"][(i * cardnum) + j]["ProductImageURL"];
            manychat["content"]["messages"][i]["elements"].push({ "title": title, "subtitle": subtitle, "image_url": image_url, "action_url": "https://manychat.com", "buttons": [] });
        }
    }
    return manychat;
}
var RedTech = /** @class */ (function () {
    function RedTech(ListName, SiteName) {
        this.SiteName = SiteName;
        this.ListName = ListName;
    }
    RedTech.generatetoken = function () {
        return __awaiter(this, void 0, void 0, function () {
            var result;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, requestp({
                            headers: { "Content-Type": "application/x-www-form-urlencoded" },
                            form: {
                                grant_type: "refresh_token",
                                client_id: this.ClientID + "@" + this.TenantID,
                                client_secret: this.ClientSecret,
                                resource: this.ApplicationID + "/" + this.TenantName + ".sharepoint.com@" + this.TenantID,
                                refresh_token: this.RefreshToken
                            },
                            uri: "https://accounts.accesscontrol.windows.net/" + this.TenantID + "/tokens/OAuth/2",
                            method: 'POST'
                        })];
                    case 1:
                        result = _a.sent();
                        return [2 /*return*/, JSON.parse(result).access_token];
                }
            });
        });
    };
    RedTech.getAccessToken = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!(this.AccessToken == null || Date.now().valueOf() > (this.IssueDate + 28500000))) return [3 /*break*/, 2];
                        console.log("Generating new token");
                        this.IssueDate = Date.now().valueOf();
                        _a = this;
                        return [4 /*yield*/, this.generatetoken()];
                    case 1:
                        _a.AccessToken = _b.sent();
                        _b.label = 2;
                    case 2: return [2 /*return*/, this.AccessToken];
                }
            });
        });
    };
    RedTech.prototype.getListItems = function () {
        return __awaiter(this, void 0, void 0, function () {
            var result, _a, _b, _c;
            var _d, _e;
            return __generator(this, function (_f) {
                switch (_f.label) {
                    case 0:
                        _a = requestp;
                        _d = {
                            json: true
                        };
                        _e = {};
                        _b = "Authorization";
                        _c = "Bearer ";
                        return [4 /*yield*/, RedTech.getAccessToken()];
                    case 1: return [4 /*yield*/, _a.apply(void 0, [(_d.headers = (_e[_b] = _c + (_f.sent()),
                                _e["Content-Type"] = "application/json; odata=verbose",
                                _e["Accept"] = "application/json; odata=nometadata",
                                _e),
                                _d.uri = encodeURI("https://" + RedTech.TenantName + ".sharepoint.com/sites/" + this.SiteName + "/_api/web/lists/GetByTitle('" + this.ListName + "')/items"),
                                _d.body = "",
                                _d.method = 'GET',
                                _d)])];
                    case 2:
                        result = _f.sent();
                        return [2 /*return*/, result];
                }
            });
        });
    };
    RedTech.TenantID = "e1cbaea0-f8c8-4ce7-a484-ade672f55f8d";
    RedTech.TenantName = "redtechly0";
    RedTech.ApplicationID = "00000003-0000-0ff1-ce00-000000000000";
    RedTech.ClientID = "bc1fdfac-50ef-4031-90c8-f12401b38e7a";
    RedTech.ClientSecret = "br+8R9XC+97p24aP8iSZaIgl16CKOxUjpcVlG41WXnM=";
    RedTech.RefreshToken = "PAQABAAEAAAD--DLA3VO7QrddgJg7WevrLVZx_-lO6VBwT_3LyQbCKCz_y31k2_NBUlGOCHVaTQNssNFOz9ciOG68LVin33nflACoDGbMXjK6GMvshqE6z7ccWyDM2Rm4Nc9WJUJazyiIxg1xCFWRNm2kwLtsw4klNSe9kEtoJQyAMTxRfqpTBn_RDXQYM7NCGkczXQ_r4V23GesvEFuggk5JDi3r9N-6Rr4C7VU7JQaX5VWn4tWDmdiGQC8d4b-lIbBMN6Iwm8dCmjQw7tenNj-K7sLKt4QioI31zIt46Q6MxsRkjUFCqCAA";
    RedTech.AccessToken = "";
    RedTech.IssueDate = 0;
    return RedTech;
}());
app.get("/DarAssendan/SharePoint/get", function (req, res) { return __awaiter(void 0, void 0, void 0, function () {
    var SiteName, ListName, darassedn, _a, _b, _c;
    return __generator(this, function (_d) {
        switch (_d.label) {
            case 0:
                SiteName = "REDCompanies";
                ListName = "DarAssendan";
                darassedn = new RedTech(ListName, SiteName);
                _b = (_a = res).send;
                _c = DarAssendansharepoint;
                return [4 /*yield*/, darassedn.getListItems()];
            case 1:
                _b.apply(_a, [_c.apply(void 0, [_d.sent()])]);
                return [2 /*return*/];
        }
    });
}); });
app.listen(PORT, function () {
    console.log("listening on post http://localhost:".concat(PORT));
});
