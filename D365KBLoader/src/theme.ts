import { createDarkTheme, createLightTheme, type BrandVariants, type Theme } from '@fluentui/react-components';

// A vivid Microsoft-style blue ramp (16 stops, required by Fluent v9).
export const kbBrand: BrandVariants = {
  10: '#020305',
  20: '#111A2E',
  30: '#162A4A',
  40: '#193863',
  50: '#1B477E',
  60: '#1B5799',
  70: '#1867B5',
  80: '#1278D2',
  90: '#2E8AE0',
  100: '#4F9DE8',
  110: '#71B0EF',
  120: '#92C3F4',
  130: '#B0D4F8',
  140: '#CCE3FB',
  150: '#E3EFFD',
  160: '#F2F8FE',
};

export const kbLightTheme: Theme = {
  ...createLightTheme(kbBrand),
  colorNeutralBackground2: '#F5F8FC',
};

export const kbDarkTheme: Theme = {
  ...createDarkTheme(kbBrand),
  colorNeutralBackground1: '#111A2E',
  colorNeutralBackground2: '#0E1525',
};

export function getHeroGradient(mode: 'light' | 'dark') {
  return mode === 'dark'
    ? 'linear-gradient(135deg, #08284C 0%, #0C4E8D 45%, #1369B8 100%)'
    : 'linear-gradient(135deg, #0B3A6F 0%, #1364B5 45%, #1F86DD 100%)';
}
