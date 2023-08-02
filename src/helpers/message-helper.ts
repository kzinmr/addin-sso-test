/* global document */

export function showMessage(text: string): void {
  document.getElementById("message-area").style.display = "flex";
  document.getElementById("message-area").innerText = text;
}

export function clearMessage(): void {
  document.getElementById("message-area").style.display = "flex";
  document.getElementById("message-area").innerText = "---<br>";
}

export function hideMessage(): void {
  document.getElementById("message-area").style.display = "none";
  document.getElementById("message-area").innerText = "---<br>";
}
