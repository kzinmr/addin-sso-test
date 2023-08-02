import { showMessage } from "./message-helper";

export async function callGetUserData(middletierToken: string): Promise<Response> {
  try {
    return await fetch("https://localhost:3000/getuserdata", {
      headers: { Authorization: `Bearer ${middletierToken}` },
    });
  } catch (err) {
    showMessage(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}
