import { sha256FontBytes, type FontSource, type FontConfigurationFragment } from '@pptx-editor.dev/core/layout';

/** Convert fixture faces into the public export font configuration. */
export function fontConfiguration(faces: readonly FontSource[]): FontConfigurationFragment {
  return { sources: faces.map((face, index) => ({
    id: `fixture-${index}`,
    request: { family: face.family, weight: face.bold ? 700 : 400, style: face.italic ? 'italic' : 'normal' },
    bytes: face.bytes,
    hash: sha256FontBytes(face.bytes),
    faceIndex: 0,
  })) };
}
