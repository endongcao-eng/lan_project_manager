document.addEventListener('click', async (event) => {
  const closeOverdue = event.target.closest('[data-close-overdue-modal]');
  if (closeOverdue) {
    const modal = closeOverdue.closest('[data-overdue-modal]');
    if (modal) modal.remove();
    return;
  }

  const confirmTarget = event.target.closest('[data-confirm]');
  if (confirmTarget && !window.confirm(confirmTarget.dataset.confirm || '确认执行该操作？')) {
    event.preventDefault();
    return;
  }

  const button = event.target.closest('[data-demo-user]');
  if (button) {
    const username = document.querySelector('input[name="username"]');
    const password = document.querySelector('input[name="password"]');
    if (username && password) {
      username.value = button.dataset.demoUser;
      password.value = button.dataset.demoPass;
      username.focus();
    }
    return;
  }

  const editorButton = event.target.closest('[data-editor-action]');
  if (editorButton) {
    const editor = document.querySelector('.document-editor');
    if (!editor || editor.readOnly) return;
    const action = editorButton.dataset.editorAction;
    const snippets = {
      'insert-heading': '\n\n## 新章节标题\n\n',
      'insert-reference': '\n\n[参考文献待补充：请使用 Zotero 插入正式引用]\n\n',
      'insert-note': '\n\n【待核查：此处需要补充可靠来源或进一步确认】\n\n'
    };
    const text = snippets[action] || '';
    const start = editor.selectionStart || 0;
    const end = editor.selectionEnd || 0;
    editor.value = editor.value.slice(0, start) + text + editor.value.slice(end);
    editor.focus();
    editor.selectionStart = editor.selectionEnd = start + text.length;
    editor.dispatchEvent(new Event('input'));
    return;
  }

  const carouselButton = event.target.closest('[data-carousel-prev], [data-carousel-next]');
  if (carouselButton) {
    const carousel = carouselButton.closest('[data-image-carousel]');
    moveCarousel(carousel, carouselButton.hasAttribute('data-carousel-next') ? 1 : -1);
    return;
  }

  const thumb = event.target.closest('[data-carousel-thumb]');
  if (thumb) {
    const carousel = thumb.closest('[data-image-carousel]');
    selectCarouselThumb(carousel, thumb);
    return;
  }

  const copyButton = event.target.closest('[data-image-copy]');
  if (copyButton) {
    await copyImage(copyButton.dataset.src, copyButton);
  }
});

function estimateMixedWordCount(text) {
  const chinese = (text.match(/[\u4e00-\u9fff]/g) || []).length;
  const english = (text.match(/[A-Za-z0-9]+(?:[-_'][A-Za-z0-9]+)*/g) || []).length;
  return chinese + english;
}

document.addEventListener('input', (event) => {
  const editor = event.target.closest('[data-count-target]');
  if (!editor) return;
  const target = document.getElementById(editor.dataset.countTarget);
  if (!target) return;
  target.textContent = `当前估算：${estimateMixedWordCount(editor.value)} 字`;
});

function selectCarouselThumb(carousel, thumb) {
  if (!carousel || !thumb) return;
  const main = carousel.querySelector('[data-carousel-main]');
  const title = carousel.querySelector('[data-carousel-title]');
  const subtitle = carousel.querySelector('[data-carousel-subtitle]');
  const open = carousel.querySelector('[data-carousel-open]');
  const copy = carousel.querySelector('[data-image-copy]');
  carousel.querySelectorAll('[data-carousel-thumb]').forEach((item) => item.classList.remove('active'));
  thumb.classList.add('active');
  if (main) main.src = thumb.dataset.src;
  if (main) main.alt = thumb.dataset.title || '';
  if (title) title.textContent = thumb.dataset.title || '';
  if (subtitle) subtitle.textContent = thumb.dataset.subtitle || '';
  if (open) open.href = thumb.dataset.src;
  if (copy) copy.dataset.src = thumb.dataset.src;
  thumb.scrollIntoView({behavior: 'smooth', inline: 'center', block: 'nearest'});
}

function moveCarousel(carousel, direction) {
  if (!carousel) return;
  const thumbs = Array.from(carousel.querySelectorAll('[data-carousel-thumb]'));
  if (!thumbs.length) return;
  const current = thumbs.findIndex((item) => item.classList.contains('active'));
  const next = (current + direction + thumbs.length) % thumbs.length;
  selectCarouselThumb(carousel, thumbs[next]);
}

async function copyImage(src, button) {
  const oldText = button.textContent;
  try {
    if (!navigator.clipboard || !window.ClipboardItem) {
      throw new Error('Clipboard image write is not supported');
    }
    const response = await fetch(src);
    const blob = await response.blob();
    await navigator.clipboard.write([new ClipboardItem({[blob.type]: blob})]);
    button.textContent = '已复制';
    setTimeout(() => { button.textContent = oldText; }, 1600);
  } catch (err) {
    window.open(src, '_blank');
    button.textContent = '已打开原图';
    setTimeout(() => { button.textContent = oldText; }, 1600);
  }
}

// 让原生日期选择框更像“点击展开日历”：支持 showPicker 的浏览器会直接弹出日历。
document.addEventListener('focusin', (event) => {
  const input = event.target.closest('input[type="date"][data-date-picker]');
  if (!input || typeof input.showPicker !== 'function') return;
  try { input.showPicker(); } catch (err) { /* 某些浏览器在非用户手势中会拒绝，保持默认行为即可 */ }
});

// 自动把普通栏目收纳为默认折叠面板，减少页面初始信息密度。
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.panel').forEach((panel) => {
    if (panel.matches('.fold-section, .overdue-modal, .emergency-banner, .task-hero, .document-editor-panel, .login-card, .login-visual')) return;
    if (panel.closest('.modal-backdrop')) return;
    const title = panel.querySelector(':scope > .section-title');
    if (!title || panel.querySelector(':scope > .panel-auto-toggle')) return;
    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'panel-auto-toggle';
    panel.insertBefore(toggle, title);
    toggle.appendChild(title);
    const body = document.createElement('div');
    body.className = 'auto-fold-body';
    while (toggle.nextSibling) body.appendChild(toggle.nextSibling);
    panel.appendChild(body);
    panel.classList.add('auto-fold-panel');
    if (!panel.querySelector('.alert')) { panel.classList.add('is-collapsed'); }
    toggle.addEventListener('click', () => panel.classList.toggle('is-collapsed'));
  });
});
