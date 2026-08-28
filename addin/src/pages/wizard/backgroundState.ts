import type { TemplateBackground, TemplateDoc } from '../../api.js';

export interface WizardBackgroundState {
  followDocument: boolean;
  customImageDataUrl: string;
}

export function defaultWizardBackgroundState(_template: Pick<TemplateDoc, 'background'> | null | undefined): WizardBackgroundState {
  return { followDocument: true, customImageDataUrl: '' };
}

export function resolveWizardBackground(
  template: Pick<TemplateDoc, 'background'> | null | undefined,
  state: WizardBackgroundState
): TemplateBackground | undefined {
  if (!state.followDocument && state.customImageDataUrl) {
    return { type: 'picture', imageDataUrl: state.customImageDataUrl };
  }
  return template?.background;
}

export function cloneTemplateWithWizardBackground(template: TemplateDoc, state: WizardBackgroundState): TemplateDoc {
  const next: TemplateDoc = JSON.parse(JSON.stringify(template));
  const background = resolveWizardBackground(template, state);
  if (background) next.background = JSON.parse(JSON.stringify(background));
  else delete next.background;
  return next;
}
