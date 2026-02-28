/**
 * camera.js - カメラ撮影・画像リサイズモジュール
 * iPhone最適化済み
 */
const CameraModule = (() => {
  const MAX_PHOTOS = 3;
  const MAX_WIDTH = 1024;
  const JPEG_QUALITY = 0.8;
  let photos = []; // base64画像の配列

  function init() {
    const addPhotoBtn = document.getElementById('addPhotoBtn');
    const photoInput = document.getElementById('photoInput');

    addPhotoBtn.addEventListener('click', () => {
      if (photos.length >= MAX_PHOTOS) {
        alert(`写真は最大${MAX_PHOTOS}枚までです`);
        return;
      }
      photoInput.click();
    });

    photoInput.addEventListener('change', handleFileSelect);
  }

  function handleFileSelect(event) {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    const remaining = MAX_PHOTOS - photos.length;
    const filesToProcess = Array.from(files).slice(0, remaining);

    filesToProcess.forEach((file) => {
      processImage(file);
    });

    // inputをリセット（同じファイルを再選択可能にする）
    event.target.value = '';
  }

  function processImage(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const resized = resizeImage(img);
        photos.push(resized);
        renderPreview();
        updateAddButton();
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  }

  function resizeImage(img) {
    const canvas = document.createElement('canvas');
    let width = img.width;
    let height = img.height;

    // 幅がMAX_WIDTHを超える場合のみリサイズ
    if (width > MAX_WIDTH) {
      height = Math.round((height * MAX_WIDTH) / width);
      width = MAX_WIDTH;
    }

    canvas.width = width;
    canvas.height = height;

    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0, width, height);

    return canvas.toDataURL('image/jpeg', JPEG_QUALITY);
  }

  function renderPreview() {
    const container = document.getElementById('photoPreview');
    container.innerHTML = '';

    photos.forEach((src, index) => {
      const item = document.createElement('div');
      item.className = 'photo-item';

      const img = document.createElement('img');
      img.src = src;
      img.alt = `写真 ${index + 1}`;

      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'photo-remove';
      removeBtn.textContent = '\u00D7';
      removeBtn.addEventListener('click', () => removePhoto(index));

      item.appendChild(img);
      item.appendChild(removeBtn);
      container.appendChild(item);
    });
  }

  function removePhoto(index) {
    photos.splice(index, 1);
    renderPreview();
    updateAddButton();
  }

  function updateAddButton() {
    const btn = document.getElementById('addPhotoBtn');
    if (photos.length >= MAX_PHOTOS) {
      btn.style.display = 'none';
    } else {
      btn.style.display = 'flex';
    }
  }

  function getPhotos() {
    return [...photos];
  }

  function clear() {
    photos = [];
    renderPreview();
    updateAddButton();
  }

  return { init, getPhotos, clear };
})();
