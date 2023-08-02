import { dialogFallback } from "./fallbackauthdialog";
import { callGetUserData } from "./middle-tier-calls";
import { showMessage } from "./message-helper";
import { handleClientSideErrors } from "./error-handler";

/* global OfficeRuntime */

let retryGetMiddletierToken = 0;

export async function getUserData(callback): Promise<void> {
  try {
    let middletierToken: string = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
    let response: any = await callGetUserData(middletierToken);
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else if (response.claims) {
      // Microsoft Graphでは、追加の認証が必要(Auth Challenge)
      // OfficeホストにClaims文字列を使用して新しいトークンを取得させ、AADに必要なすべての認証フォームをユーザーに求めるように指示します。
      let mfaMiddletierToken: string = await OfficeRuntime.auth.getAccessToken({
        authChallenge: response.claims,
      });
      response = callGetUserData(mfaMiddletierToken);
    }
    // AADエラーはHTTPコード200でクライアントに返されるため、以下のcatchブロックはトリガーされない。
    if (response.error) {
      handleAADErrors(response, callback);
    } else {
      callback(response);
    }
  } catch (exception) {
    // もしhandleClientSideErrorsがtrueを返したら、フォールバックダイアログで認証を試みる。
    // Officeにサインインしていない、同意プロンプトを中止した、等の原因.
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log("Path1: ", exception.code);
        dialogFallback(callback);
      }
    } else {
      showMessage("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}

function handleAADErrors(response: any, callback: any): void {
  // まれにミドル・ティア・トークンが、Officeが検証する時点では未失効だが、AADに送信されるまでに失効することがあるケースに対応する。
  // これに対しAADは "提供された'assertion'の値は有効ではありません。assertionは期限切れです" と応答する。
  // getAccessToken の呼び出しを1回だけ再試行します。このときOfficeは、有効期限が切れていない新しいミドルティアトークンを返します。
  if (response.error_description.indexOf("AADSTS500133") !== -1 && retryGetMiddletierToken <= 0) {
    retryGetMiddletierToken++;
    getUserData(callback);
  } else {
    console.log("Path2:");
    dialogFallback(callback);
  }
}
