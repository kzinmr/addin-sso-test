/* global document, Office, Word */

// import { getUserData } from "../helpers/sso-helper";
import { dialogFallback } from "../helpers/fallbackauthdialog";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("getProfileButton").onclick = run;
  }
});

export async function run() {
  dialogFallback(writeDataToOfficeDocument);
}

export function writeDataToOfficeDocument(obj: Object): Promise<any> {
  return Word.run((context) => {
    let data: string[] = [];
    let userProfileInfo: string[] = [];
    userProfileInfo.push(obj["displayName"]);
    userProfileInfo.push(obj["mail"]);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        data.push(userProfileInfo[i]);
      }
    }

    const documentBody: Word.Body = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
