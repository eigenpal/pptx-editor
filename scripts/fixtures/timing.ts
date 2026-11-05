// Slide timing XML for tests (ECMA-376 Part 1, §19.5 `p:timing`): a root time node, a main
// click sequence of entrance effects, interactive sequences, and build entries. Each writer
// numbers its `p:cTn/@id` values from 10 on; the root is 1 and the main sequence is 2.

export interface TimingWriter {
  /** One behavior that shows shape `spid` (`p:set` of `style.visibility`). */
  show(spid: number): string;
  /** One entrance effect on `spid`; `extra` adds behaviors; `id` fixes its `p:cTn/@id`. */
  effect(spid: number, extra?: string, id?: number): string;
  /** One click step of the main sequence holding `effects`. */
  click(effects: string): string;
  /** The main sequence holding `steps`. */
  mainSeq(steps: string): string;
  /** An interactive sequence: a click on shape `trigger` plays an entrance of `spid`. */
  interactive(trigger: number, spid: number): string;
  /** `p:timing` with `sequences` under the root, and one `p:bldP` for each id of `builds`. */
  timing(sequences: string, builds: readonly number[]): string;
}

export function timingWriter(): TimingWriter {
  let next = 10;
  const w: TimingWriter = {
    show: (spid) =>
      `<p:set><p:cBhvr><p:cTn id="${next++}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
      `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
      '<p:to><p:strVal val="visible"/></p:to></p:set>',
    effect: (spid, extra = '', id = next++) =>
      `<p:par><p:cTn id="${id}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">` +
      `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${w.show(spid)}${extra}</p:childTnLst></p:cTn></p:par>`,
    click: (effects) =>
      `<p:par><p:cTn id="${next++}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn id="${next++}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par>` +
      '</p:childTnLst></p:cTn></p:par>',
    mainSeq: (steps) =>
      `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${steps}</p:childTnLst></p:cTn>` +
      '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
      '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>',
    interactive: (trigger, spid) =>
      `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${next++}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
      `<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${trigger}"/></p:tgtEl></p:cond></p:stCondLst>` +
      `<p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>${w.effect(spid)}</p:childTnLst></p:cTn>` +
      `<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${trigger}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`,
    timing: (sequences, builds) =>
      '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
      sequences +
      '</p:childTnLst></p:cTn></p:par></p:tnLst>' +
      (builds.length
        ? `<p:bldLst>${builds.map((id) => `<p:bldP spid="${id}" grpId="0" animBg="1"/>`).join('')}</p:bldLst>`
        : '') +
      '</p:timing>',
  };
  return w;
}

/** One click entrance for each shape of `spids`, in order, with a build entry for each. */
export function clickEntrances(spids: readonly number[]): string {
  const w = timingWriter();
  return w.timing(w.mainSeq(spids.map((id) => w.click(w.effect(id))).join('')), spids);
}
