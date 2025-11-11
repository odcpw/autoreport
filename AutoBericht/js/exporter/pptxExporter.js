import { readFileAsDataUrl } from '../utils/fileUtils.js';

// NOTE: PPT export logic stays here as scaffolding only; revisit once presentation work resumes.

const MAX_IMAGE_EDGE = 1600;
const IMAGE_QUALITY = 0.85;

export async function exportReportPpt(snapshot) {
  ensureLibrary();
  const pptx = new window.PptxGenJS();
  const filename = buildFilename(snapshot, 'Report Presentation');
  const chapters = snapshot?.project?.report?.chapters || [];
  const photos = snapshot?.project?.photos || {};

  if (!chapters.length) {
    const slide = pptx.addSlide();
    slide.addText('No chapters available for export.', { x: 1, y: 1, fontSize: 18 });
  } else {
    for (const chapter of chapters.filter((item) => item.includeInReport !== false)) {
      const slide = pptx.addSlide();
      slide.addText(`${chapter.id || ''} — ${chapter.title || ''}`, {
        x: 0.5,
        y: 0.5,
        fontSize: 24,
        bold: true,
      });
      const text = chapter.finding?.useAdjusted
        ? chapter.finding.adjustedText || chapter.finding.masterText
        : chapter.finding?.masterText;
      slide.addText(text || 'No finding text provided.', {
        x: 0.5,
        y: 1.2,
        fontSize: 16,
        color: '363636',
        w: 9,
        h: 2,
      });

      const recommendations = Object.keys(chapter.recommendations || {})
        .map((key) => {
          const rec = chapter.recommendations[key];
          const textRec = rec?.useAdjusted ? rec.adjusted || rec.master : rec?.master;
          return textRec ? `• ${textRec}` : '';
        })
        .filter(Boolean)
        .join('\n');

      if (recommendations) {
        slide.addText(recommendations, {
          x: 0.5,
          y: 3.3,
          fontSize: 14,
          color: '2d2d2d',
          w: 9,
          h: 2.5,
          bullet: true,
        });
      }

      const chapterPhotos = collectPhotosForChapter(snapshot.project, chapter.id);
      if (chapterPhotos.length) {
        const grid = buildPhotoGrid(chapterPhotos.slice(0, 6));
        await embedPhotoGrid(slide, grid, photos);
        if (chapterPhotos.length > 6) {
          slide.addText(`+ ${chapterPhotos.length - 6} more`, {
            x: 0.5,
            y: 5.8,
            fontSize: 12,
            color: '777777',
          });
        }
      }
    }
  }

  await pptx.writeFile({ fileName: filename });
}

export async function exportSeminarPpt(snapshot) {
  ensureLibrary();
  const pptx = new window.PptxGenJS();
  const filename = buildFilename(snapshot, 'Seminar Slides');
  const photos = snapshot?.project?.photos || {};

  const seminarGroups = {};
  Object.entries(photos).forEach(([path, info]) => {
    (info.tags?.seminar || []).forEach((seminarTag) => {
      if (!seminarGroups[seminarTag]) seminarGroups[seminarTag] = [];
      seminarGroups[seminarTag].push({ path, info });
    });
  });

  const tags = Object.keys(seminarGroups);
  if (!tags.length) {
    const slide = pptx.addSlide();
    slide.addText('No seminar-tagged photos available.', { x: 1, y: 1, fontSize: 18 });
  } else {
    for (const tag of tags) {
      const slide = pptx.addSlide();
      slide.addText(`Seminar: ${tag}`, { x: 0.5, y: 0.5, fontSize: 24, bold: true });

      seminarGroups[tag]
        .slice(0, 6)
        .forEach((photo, index) => {
          slide.addText(`${index + 1}. ${photo.path}`, {
            x: 0.5,
            y: 1.2 + index * 0.6,
            fontSize: 14,
            w: 4.5,
            h: 0.5,
          });
          if (photo.info?.notes) {
            slide.addText(photo.info.notes, {
              x: 5.1,
              y: 1.2 + index * 0.6,
              fontSize: 12,
              color: '555555',
              w: 4,
              h: 0.5,
            });
          }
        });

      const grid = buildPhotoGrid(seminarGroups[tag].slice(0, 6));
      await embedPhotoGrid(slide, grid, photos);

      if (seminarGroups[tag].length > 6) {
        slide.addText(`+ ${seminarGroups[tag].length - 6} more`, {
          x: 0.5,
          y: 5.8,
          fontSize: 12,
          color: '777777',
        });
      }
    }
  }

  await pptx.writeFile({ fileName: filename });
}

function ensureLibrary() {
  if (!window.PptxGenJS) {
    throw new Error('PptxGenJS library not loaded. Ensure libs/pptxgenjs/pptxgen.min.js is available.');
  }
}

function buildFilename(snapshot, suffix) {
  const today = new Date().toISOString().slice(0, 10);
  const company = snapshot?.project?.meta?.company || 'Company';
  return `${today} ${company} ${suffix}.pptx`;
}

function collectPhotosForChapter(project, chapterId) {
  if (!chapterId || !project?.photos) return [];
  const matches = [];
  Object.entries(project.photos).forEach(([path, info]) => {
    if (info.tags?.chapters?.includes?.(chapterId)) {
      matches.push({ path, info });
    }
  });
  return matches;
}

function buildPhotoGrid(photos) {
  const positions = [
    { x: 0.5, y: 5.2 },
    { x: 3.5, y: 5.2 },
    { x: 6.5, y: 5.2 },
    { x: 0.5, y: 7 },
    { x: 3.5, y: 7 },
    { x: 6.5, y: 7 },
  ];
  const size = { w: 2.5, h: 1.5 };
  return photos.map((photo, index) => ({
    photo,
    x: positions[index]?.x ?? 0.5,
    y: positions[index]?.y ?? 5.2,
    w: size.w,
    h: size.h,
  }));
}

async function embedPhotoGrid(slide, grid, allPhotos) {
  const tasks = grid.map((entry) => addImageToSlide(slide, entry, allPhotos));
  await Promise.all(tasks);
}

async function addImageToSlide(slide, entry, allPhotos) {
  const { photo, x, y, w, h } = entry;
  const source = allPhotos?.[photo.path]?.file;
  const captionOptions = {
    x,
    y: y + h + 0.05,
    w,
    h: 0.3,
    fontSize: 10,
    color: '555555',
  };

  if (source instanceof File) {
    try {
      const dataUrl = await loadAndScaleImage(source);
      if (dataUrl) {
        slide.addImage({ data: dataUrl, x, y, w, h });
        const caption = buildCaption(photo);
        slide.addText(caption, captionOptions);
        return;
      }
    } catch (error) {
      console.warn('Failed to embed image', error);
    }
  }

  slide.addShape(window.PptxGenJS.ShapeType.rect, {
    x,
    y,
    w,
    h,
    fill: { color: 'DDDDDD' },
    line: { color: '999999' },
  });
  slide.addText(buildCaption(photo), captionOptions);
}

function buildCaption(photo) {
  const parts = [photo.path];
  if (photo.info?.notes) {
    parts.push(photo.info.notes.replace(/\s+/g, ' '));
  }
  return parts.join(' — ');
}

async function loadAndScaleImage(file) {
  const dataUrl = await readFileAsDataUrl(file);
  const image = await loadImage(dataUrl);
  const { width, height } = scaleDimensions(image.width, image.height, MAX_IMAGE_EDGE);
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(image, 0, 0, width, height);
  return canvas.toDataURL('image/jpeg', IMAGE_QUALITY);
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });
}

function scaleDimensions(width, height, maxEdge) {
  if (!width || !height) {
    return { width: maxEdge, height: Math.round(maxEdge * 0.75) };
  }
  const largest = Math.max(width, height);
  if (largest <= maxEdge) {
    return { width, height };
  }
  const scale = maxEdge / largest;
  return { width: Math.round(width * scale), height: Math.round(height * scale) };
}
