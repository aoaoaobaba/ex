document.addEventListener("DOMContentLoaded", () => {
    const targetButtons = ["disabledButton1", "disabledButton2"];
  
    // 無効なボタンを有効化 & 赤字にする
    targetButtons.forEach((name) => {
      const button = document.querySelector(`button[name='${name}']`);
      if (button && button.disabled) {
        button.disabled = false;
        button.style.color = "red";
      }
    });
  });
  