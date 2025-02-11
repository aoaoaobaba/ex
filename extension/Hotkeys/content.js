document.addEventListener("DOMContentLoaded", () => {
    const urlPrefix1 = "https://example.com/specific-page"; // 既存の対象URL
    const urlPrefix2 = "https://example.com/another-page"; // 新しい対象URL
    const currentUrl = window.location.href;
  
    if (currentUrl.startsWith(urlPrefix1)) {
      // URL1 の機能: ボタンのキャプション変更とショートカット操作
      const buttonMappings = {
        "buttonName1": "A",
        "buttonName2": "B",
        "buttonName3": "C"
      };
  
      Object.entries(buttonMappings).forEach(([name, key]) => {
        const button = document.querySelector(`button[name='${name}']`);
        if (button) {
          button.innerHTML += ` <span style='color: red;'>(${key})</span>`;
        }
      });
  
      document.addEventListener("keydown", async (e) => {
        if (e.ctrlKey && e.shiftKey) {
          const pressedKey = e.key.toUpperCase();
          const targetButton = Object.entries(buttonMappings).find(
            ([, key]) => key === pressedKey
          );
          if (targetButton) {
            document.querySelector(`button[name='${targetButton[0]}']`)?.click();
          }
          if (pressedKey === "V") {
            try {
              const text = await navigator.clipboard.readText();
              const values = text.split("\n");
              const inputFields = [
                "inputName1",
                "inputName2",
                "inputName3",
                "inputName4"
              ];
              inputFields.forEach((name, index) => {
                const field = document.querySelector(`input[name='${name}']`);
                if (field) field.value = values[index] || "";
              });
              document.querySelector("button[name='btnRegist']")?.focus();
            } catch (err) {
              console.error("Clipboard access error", err);
            }
          }
        }
      });
    }
  
    if (currentUrl.startsWith(urlPrefix2)) {
      // URL2 の機能: 無効なボタンを有効化し、赤字にする
      const targetButtons = ["disabledButton1", "disabledButton2"];
      targetButtons.forEach((name) => {
        const button = document.querySelector(`button[name='${name}']`);
        if (button && button.disabled) {
          button.disabled = false;
          button.style.color = "red";
        }
      });
    }
  });
  