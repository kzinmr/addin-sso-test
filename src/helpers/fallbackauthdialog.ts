/* global console, localStorage, Office, window */

import { Configuration, LogLevel, PublicClientApplication, RedirectRequest } from "@azure/msal-browser";
import { callGetUserData } from "./middle-tier-calls";
import { showMessage } from "./message-helper";

// Provider Info registered on Azure AD
// See. https://learn.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2
const clientId = "YOUR_CLIENT_ID_HERE";
const authority = "https://login.microsoftonline.com/common";
const redirect = `https://${window.location.host}/fallbackauthdialog.html`;
const exposedAPI = `api://${window.location.host}/${clientId}/access_as_user`;
const loginRequest: RedirectRequest = { scopes: [exposedAPI], extraScopesToConsent: ["user.read"] };
const msalConfig: Configuration = {
  auth: {
    clientId: clientId,
    authority: authority,
    redirectUri: redirect,
    navigateToLoginRequestUrl: true,
  },
  cache: {
    cacheLocation: "localStorage", // Needed to avoid "User login is required" error.
    storeAuthStateInCookie: true, // Recommended to avoid certain IE/Edge issues.
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
    },
  },
};
const publicClientApp: PublicClientApplication = new PublicClientApplication(msalConfig);

Office.onReady(() => {
  // messageParent はdialog popupが開いたときに有効化
  if (Office.context.ui.messageParent) {
    publicClientApp
      .handleRedirectPromise() // リダイレクトによる認証/トークン取得をハンドリング
      .then((tokenResponse) => {
        if (tokenResponse.tokenType === "id_token") {
          // assumed to jump here from `loginRedirect`
          localStorage.setItem("loggedIn", "yes");
        } else {
          // asumed to jump here from `acquireTokenRedirect`
          // 取得したアクセストークンのtaskpaneへの送信 (messageParent)
          Office.context.ui.messageParent(
            JSON.stringify({
              status: "success",
              result: tokenResponse.accessToken,
              accountId: tokenResponse.account.homeAccountId,
            })
          );
        }
      })
      .catch((error) => {
        console.log(error);
        Office.context.ui.messageParent(
          JSON.stringify({
            status: "failure",
            result: error,
          })
        );
      });

    if (localStorage.getItem("loggedIn") !== "yes") {
      // loginRedirect により
      // - ユーザがログインし
      // - handleRedirectPromise のthenコールバックが実行され
      // - localStorage.setItem("loggedIn", "yes") され
      // - ダイアログ内でリダイレクトされ、今度は acquireTokenRedirect が実行される
      publicClientApp.loginRedirect(loginRequest);
    } else {
      // 初回ログイン時にlocalStorageに値がない状態で acquireTokenRedirect を呼ぶと
      // "User login is required" というエラーが発生するため、まずloginRedirectを呼ぶ
      publicClientApp.acquireTokenRedirect(loginRequest);
    }
  }
});

let loginDialog: Office.Dialog = null;
let callbackFunction = null;
export async function dialogFallback(callback) {
  // ユーザが既にサインインしている場合、トークンの取得をサイレントに試みる
  if (publicClientApp.getActiveAccount() !== null) {
    console.log("active accont:", publicClientApp.getActiveAccount());
    const result = await publicClientApp.acquireTokenSilent(loginRequest);
    if (result !== null && result.accessToken !== null) {
      const response = await callGetUserData(result.accessToken);
      const json = await response.json();
      console.log("Response from /getuserdata:", JSON.stringify(json));
      callback(json);
    }
  } else {
    // イベントハンドラ内でも使用するため、グローバル変数に格納
    callbackFunction = callback;
    // サインインしてない場合、Office dialog API で認証ダイアログポップアップを開く
    Office.context.ui.displayDialogAsync(redirect, { height: 60, width: 30 }, (result) => {
      console.log("Dialog has initialized. Wiring up events");
      loginDialog = result.value;
      loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    });
  }
}

// 認証ダイアログに登録されるイベントハンドラ
// 引数はダイアログポップアップからリダイレクトで渡される (See. handleRedirectPromise)
async function processMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message);

  if (messageFromDialog.status === "success") {
    loginDialog.close();
    const accountId = messageFromDialog.accountId;
    const middletierToken = messageFromDialog.result;

    // サインインしたアカウントを、今後のリクエストのアクティブアカウントとして使用するようにMSALを設定する。
    const homeAccount = publicClientApp.getAccountByHomeId(accountId);
    if (homeAccount) {
      publicClientApp.setActiveAccount(homeAccount);
    }
    // トークンを使って、ユーザーデータを取得
    const response = await callGetUserData(middletierToken);
    console.log("Response from /getuserdata:", JSON.stringify(await response.json()));
    callbackFunction(response);
  } else {
    // Something went wrong with auth(n/z) of the web application.
    // loginDialog.close();

    if (messageFromDialog.error) {
      showMessage(JSON.stringify(messageFromDialog.error.toString()));
    } else {
      console.error(messageFromDialog.result);
    }
  }
}
