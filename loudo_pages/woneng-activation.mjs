export function needsBrowserHelp(userAgent, mode) {
  return mode === 'scheme' && /MicroMessenger/i.test(userAgent);
}

export function installLaunchHelp(doc, userAgent) {
  const button = doc.getElementById('open-mini');
  const help = doc.getElementById('launch-help');
  if (!button || !help) return;
  if (needsBrowserHelp(userAgent, button.dataset.mode)) {
    help.hidden = false;
    help.textContent = '你正在微信内浏览。请点右上角「···」，选择在浏览器中打开，再点击上方按钮；也可以在微信搜索「我能的」→ 卡密激活。';
    button.addEventListener('click', (event) => {
      event.preventDefault();
      help.focus();
    });
  } else {
    button.addEventListener('click', () => {
      help.hidden = false;
      help.textContent = '如出现系统提示，请允许打开微信。若未跳转，请查看下方帮助；此页面不会自动激活或扣费。';
    });
  }
}
if (typeof document !== 'undefined') installLaunchHelp(document, navigator.userAgent);
