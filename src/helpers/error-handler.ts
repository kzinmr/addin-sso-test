import { showMessage } from "./message-helper";

export function handleClientSideErrors(error: any): boolean {
  let invokeFallBackDialog: boolean = false;
  switch (error.code) {
    case 13001:
      // 誰もOfficeにサインインしていない。
      // Officeログイン無しでアドインを効果的に使用できない場合は、 getAccessTokenの最初の呼び出しで`allowSignInPrompt: true`オプションを渡す必要
      //
      showMessage(
        "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
      );
      return invokeFallBackDialog;
    case 13002:
      // ユーザーは同意プロンプトを中止した。
      // 同意無しにアドインを効果的に使用できない場合は、getAccessTokenの最初の呼び出しで`allowConsentPrompt: true`オプションを渡すべき
      showMessage(
        "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
      );
      return invokeFallBackDialog;
    case 13006:
      // Only seen in Office on the Web.
      showMessage(
        "Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
      );
      return invokeFallBackDialog;
    case 13008:
      // Only seen in Office on the Web.
      showMessage("Office is still working on the last operation. When it completes, try this operation again.");
      return invokeFallBackDialog;
    case 13010:
      // Only seen in Office on the Web.
      showMessage("Follow the instructions to change your browser's zone configuration.");
      return invokeFallBackDialog;
    default:
      // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001,
      // fall back to non-SSO sign-in.
      invokeFallBackDialog = true;
      return invokeFallBackDialog;
  }
}
