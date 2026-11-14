interface TextModel { ps: unknown[] }
interface CellModel { tx: TextModel; fill?: unknown; bd: Record<string, unknown>; anchor?: string }
interface ShapeModel { ph?: { type?: string }; tx: TextModel; tbl?: { rows: Array<{ cells: CellModel[]; h?: number }>; cols?: number[] }; x: number; y: number; w: number; h: number }
interface SlideModel { shapes: ShapeModel[]; trans?: Record<string, unknown> }
interface PresentationModel { designs: Record<string, unknown>; slides: SlideModel[]; props: Record<string, unknown>; hf: Record<string, unknown>; W: number; H: number }
interface TemplateRuntime {
  L?: {
    model?: {
      newPresentation(options: Record<string, unknown>): PresentationModel;
      newSlide(presentation: PresentationModel, layout: string, designId: string): SlideModel;
      newTable(presentation: PresentationModel, slide: SlideModel, rows: number, columns: number, x: number, y: number, width: number): ShapeModel;
    };
    txt?: { para(text: string, paragraph?: Record<string, unknown>, run?: Record<string, unknown>): unknown; setPlain(text: TextModel, value: string): void; fromLines(lines: string[]): unknown[] };
    diagram?: { make(presentation: PresentationModel, slide: SlideModel, shape: ShapeModel | undefined, layout: string, lines: string[], options?: Record<string, unknown>): unknown };
    clone?<T>(value: T): T;
    templates?: { list(): TemplateInfo[]; build(id: string, size?: Record<string, number>): PresentationModel | null };
  };
}
interface TemplateInfo { id: string; name: string; design: string }

const templates: TemplateInfo[] = [
  { id: 'status', name: 'Project Status Report', design: 'blueprint' },
  { id: 'meeting', name: 'Team Meeting', design: 'meadow' },
  { id: 'lesson', name: 'Lesson', design: 'parchment' },
  { id: 'event', name: 'Event Announcement', design: 'night' },
];

/** Install the PPTX content-template service after the document model has loaded. */
export function installPptxTemplates(runtime: Window): void {
  const L = (runtime as Window & TemplateRuntime).L;
  const model = L?.model;
  const text = L?.txt;
  const diagram = L?.diagram;
  if (!L || !model || !text || !diagram) throw new Error('PPTX model runtime is not ready for content templates');

  const addTable = (presentation: PresentationModel, slide: SlideModel, placeholder: ShapeModel, rows: string[][], widths: number[]) => {
    const table = model.newTable(presentation, slide, rows.length, rows[0]?.length || 0, placeholder.x, placeholder.y, placeholder.w);
    const properties = placeholder as ShapeModel & { ph: Record<string, unknown> };
    const target = table as ShapeModel & { ph?: Record<string, unknown>; tbl: NonNullable<ShapeModel['tbl']> };
    target.ph = L.clone ? L.clone(properties.ph) : { ...properties.ph };
    const rowHeight = Math.min(44, placeholder.h / rows.length);
    rows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
      const cell = target.tbl.rows[rowIndex]?.cells[columnIndex];
      if (!cell) return;
      cell.tx.ps = [text.para(value, { algn: 'l' }, rowIndex === 0 ? { sz: 18, b: true, color: 'lt1' } : { sz: 16 })];
      cell.fill = rowIndex === 0 ? { t: 'solid', c: 'accent1', a: 1 } : { t: 'solid', c: rowIndex % 2 ? '#F2F5FA' : '#FFFFFF', a: 1 };
      for (const edge of ['l', 'r', 't', 'b']) cell.bd[edge] = { c: '#B4C3DC', w: 1, dash: 'solid' };
      cell.anchor = 'ctr';
    }));
    target.tbl.rows.forEach((row) => { row.h = rowHeight; });
    target.h = rowHeight * rows.length;
    target.tbl.cols = widths.map((fraction) => placeholder.w * fraction);
    slide.shapes.splice(slide.shapes.indexOf(placeholder), 1, target);
    return target;
  };
  const addDiagram = (presentation: PresentationModel, slide: SlideModel, layout: string, lines: string[]) => diagram.make(presentation, slide, slide.shapes.find((shape) => shape.ph?.type === 'body'), layout, lines);

  const builders: Record<string, (add: (layout: string, title: string, ...columns: string[][]) => SlideModel, presentation: PresentationModel) => void> = {
    status(add, presentation) {
      add('title', 'Project Name', ['Status report · Reporting period', 'Presented by your name']);
      add('text', 'Summary', ['Overall status: On track / At risk / Off track', 'One sentence on progress since the last report', 'One sentence on what needs attention']);
      addDiagram(presentation, add('text', 'Project Phases'), 'chevron', ['Discover', 'Design', 'Build', 'Launch']);
      const slide = add('table', 'Milestones');
      const placeholder = slide.shapes.find((shape) => shape.ph?.type === 'tbl');
      if (placeholder) addTable(presentation, slide, placeholder, [['Milestone', 'Owner', 'Due', 'Status'], ['Requirements signed off', 'Name', 'Date', 'Done'], ['Design complete', 'Name', 'Date', 'On track'], ['First release', 'Name', 'Date', 'At risk'], ['Launch', 'Name', 'Date', 'Not started']], [0.42, 0.2, 0.16, 0.22]);
      add('twoText', 'Done and Next', ['Done this period', '\tWhat was finished', '\tWhat was delivered'], ['Next period', '\tWhat comes next', '\tWho does it']);
      add('text', 'Risks and Issues', ['Risk: what could go wrong', '\tImpact and what we are doing about it', 'Issue: what is already wrong', '\tWhat it blocks and who is fixing it']);
      add('text', 'Decisions Needed', ['What you need from the audience', 'By when, and what happens if it slips']);
    },
    meeting(add, presentation) {
      add('title', 'Team Meeting', ['Date · Time · Room or video link']);
      addDiagram(presentation, add('text', 'Agenda'), 'vProcess', ['Welcome and check-in', 'Updates from last week', 'Topic for discussion', 'Action items and next steps']);
      add('twoText', 'Updates', ['Wins', '\tWhat went well', '\tWho to thank'], ['Blockers', '\tWhat is in the way', '\tWhat help is needed']);
      add('text', 'Discussion', ['The question to decide today', '\tOption A', '\tOption B']);
      add('text', 'Action Items', ['Action — owner — due date', 'Action — owner — due date', 'Action — owner — due date']);
      add('title', 'Thank You', ['Next meeting: date and time']);
    },
    lesson(add, presentation) {
      add('title', 'Lesson Title', ['Course name · Teacher name']);
      add('text', 'Today We Will', ['Learn what …', 'Practise how to …', 'Be able to explain …']);
      addDiagram(presentation, add('text', 'Key Idea'), 'radial', ['The idea', '\tExample', '\tWhy it matters', '\tCounter-example', '\tWhere it is used']);
      add('twoText', 'Compare', ['First thing', '\tFeature', '\tFeature'], ['Second thing', '\tFeature', '\tFeature']);
      add('text', 'Your Turn', ['A question or exercise for the class', 'How long they have, and in pairs or alone']);
      add('text', 'Summary and Homework', ['Three things to remember', 'Homework: what to do, due when']);
    },
    event(add, presentation) {
      add('title', 'Event Name', ['Day, date · Time', 'Place']);
      add('text', 'What to Expect', ['The highlight of the event', 'Who is speaking or performing', 'Food, drinks and activities']);
      addDiagram(presentation, add('text', 'Programme'), 'timeline', ['Time — opening', 'Time — main part', 'Time — break', 'Time — closing']);
      add('text', 'Getting There', ['Address', 'Public transport and parking', 'Accessibility']);
      add('title', 'See You There', ['RSVP by date · contact']);
    },
  };

  L.templates = {
    list: () => templates,
    build(id, size) {
      const template = templates.find((entry) => entry.id === id);
      if (!template) return null;
      const presentation = model.newPresentation({ design: template.design, empty: true, ...(size || {}) });
      const designId = Object.keys(presentation.designs)[0];
      if (!designId) throw new Error(`PPTX template ${id} has no design`);
      const add = (layout: string, title: string, ...columns: string[][]) => {
        const slide = model.newSlide(presentation, layout, designId);
        presentation.slides.push(slide);
        const placeholders = (types: string[]) => slide.shapes.filter((shape) => shape.ph && types.includes(shape.ph.type || ''));
        const titleShape = placeholders(['ctrTitle', 'title'])[0];
        if (titleShape) text.setPlain(titleShape.tx, title);
        if (layout === 'title') { const subtitle = placeholders(['subTitle'])[0]; if (subtitle && columns[0]) subtitle.tx.ps = columns[0].map((line) => text.para(line)); }
        else placeholders(['body']).forEach((body, index) => { if (columns[index]) body.tx.ps = text.fromLines(columns[index]); });
        slide.trans = { type: 'fade', spd: 'med', click: true, after: null };
        return slide;
      };
      const build = builders[id];
      if (!build) return null;
      build(add, presentation);
      presentation.props.title = '';
      Object.assign(presentation.hf, { num: true, notOnTitle: true });
      return presentation;
    },
  };
}
