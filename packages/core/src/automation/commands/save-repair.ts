// The package as every save writes it: tables with their merge markers (`table-repair.ts`) and
// the comment structure joined (`comment-repair.ts`). The editor, the automation hosts, and a
// slide copy save through it, so they write one file for one package.

import type { PptxPackage } from '../../store/index.ts';
import { repairCommentsForSave } from './comment-repair.ts';
import { repairTablesForSave } from './table-repair.ts';

/** The package as a save writes it. The store does not change. */
export function packageForSave(pkg: PptxPackage): PptxPackage {
  return repairCommentsForSave(repairTablesForSave(pkg));
}
