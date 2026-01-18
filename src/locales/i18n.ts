import i18n from 'i18next';
import { initReactI18next } from 'react-i18next';
import LanguageDetector from 'i18next-browser-languagedetector';
import { bitable } from '@lark-base-open/js-sdk';
import translationEN from './en.json';
import translationZH from './zh.json';
import translationJA from './ja.json';

const resources = {
  'zh-CN': {
    translation: translationZH,
  },
  'en-US': {
    translation: translationEN,
  },
  'ja-JP': {
    translation: translationJA,
  },
};

const normalizeLanguage = (lng: string) => {
  const lower = lng.toLowerCase();
  if (lower.startsWith('zh')) {
    return 'zh-CN';
  }
  if (lower.startsWith('ja')) {
    return 'ja-JP';
  }
  return 'en-US';
};

i18n
  .use(LanguageDetector)
  .use(initReactI18next)
  .init({
    resources,
    fallbackLng: 'en-US',
    interpolation: {
      escapeValue: false,
    },
  });

const syncLanguage = async () => {
  try {
    if (!bitable?.bridge?.getLanguage) {
      return;
    }
    const lng = await bitable.bridge.getLanguage();
    const normalized = normalizeLanguage(lng);
    if (i18n.language !== normalized) {
      i18n.changeLanguage(normalized);
    }
  } catch {
    // Ignore when running outside Base.
  }
};

syncLanguage();

export default i18n;
