document.addEventListener("DOMContentLoaded", () => {
  const buttonMappings = {
    "buttonName1": "A",
    "buttonName2": "B",
    "buttonName3": "C"
  };

  // ボタンのキャプション変更
  Object.entries(buttonMappings).forEach(([name, key]) => {
    const button = document.querySelector(`button[name='${name}']`);
    if (button) {
      button.innerHTML += ` <span style='color: red;'>(${key})</span>`;
    }
  });

  // ショートカットキー処理
  document.addEventListener("keydown", async (e) => {
    if (e.ctrlKey && e.shiftKey) {
      const pressedKey = e.key.toUpperCase();

      // 指定されたボタンをクリック
      const targetButton = Object.entries(buttonMappings).find(
        ([, key]) => key === pressedKey
      );
      if (targetButton) {
        document.querySelector(`button[name='${targetButton[0]}']`)?.click();
      }

      // Ctrl+Shift+V でクリップボードからデータを貼り付け
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

          // 登録ボタンにフォーカス
          document.querySelector("button[name='btnRegist']")?.focus();
        } catch (err) {
          console.error("Clipboard access error", err);
        }
      }
    }
  });
});
