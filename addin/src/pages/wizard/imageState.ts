import type { ImageResult } from '../../api.js';

export interface ImgSlotState {
  q: string;
  page: number;
  images: ImageResult[];
  pageSize: number;
  selected: string;
  providerError: string;
}

const DEFAULT_IMAGE_PAGE_SIZE = 9;
const MIN_IMAGE_PAGE_SIZE = 6;
const MAX_IMAGE_PAGE_SIZE = 14;

export function normalizeImagePageSize(value: unknown, fallback = DEFAULT_IMAGE_PAGE_SIZE): number {
  const n = Number(value);
  return Number.isFinite(n) && n >= MIN_IMAGE_PAGE_SIZE && n <= MAX_IMAGE_PAGE_SIZE
    ? n
    : fallback;
}

export function defaultImgState(pageSize = DEFAULT_IMAGE_PAGE_SIZE): ImgSlotState {
  return { q: '', page: 1, images: [], pageSize: normalizeImagePageSize(pageSize), selected: '', providerError: '' };
}

export function thumbnailUrlOf(img: Pick<ImageResult, 'imageUrl' | 'thumbnailUrl'>): string {
  return img.thumbnailUrl || img.imageUrl || '';
}
