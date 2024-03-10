function onOpen(e) {
  DocumentApp.getUi().createMenu("図表番号")
    .addItem("自動採番", "autoLabeling")
    .addToUi()
}

function autoLabeling() {
  const body = DocumentApp.getActiveDocument().getBody()
    console.log(body)

  const stm = new StateMachine()
  const visitor = new Visitor(element => {
    if (element.getType && element.getType() === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
      return false
    }
    stm.transition(element)
    if (!isContainer(element)) {
      return false
    }
    return true
  })
  visitor.visit(body)
}

function typeGuard(element, type) {
  if (element.getType() !== type) {
    throw new Error(`Unexpected type: ${element.getType().toString()}, ${element.getText()}`)
  }
}

class Label {
  static labeling(chapter, numberingCounter, variable) {
    const chapterNumber = new ChapterNumber(chapter)
    var labelType = '表'
    if (Label.isFigure(variable)) {
      labelType = '図'
    }

    var label = labelType
    if (!chapter) {
      label = `${labelType}${numberingCounter.numberingFor(variable)}`
    }
    else {
      label = `${labelType}${chapterNumber.number}-${numberingCounter.numberingFor(variable)}`
    }
    variable.replaceText('.*', label)
  }

  static isFigure(variable) {
    const match = variable.getText().match(/^(図|figure|fig)/i)
    return !!match
  }

  static isTable(variable) {
    const match = variable.getText().match(/^(表|table|tbl)/i)
    return !!match
  }
}

class ChapterNumber {
  constructor(chapter) {
    const format = /^(?<chapterNumber>\d+)[.]?/
    const matched = chapter.getText().match(format)
    this.number = 1
    if (!matched || !matched.groups.chapterNumber) {
      return
    }
    this.number = Number(matched.groups.chapterNumber)
  }
}

class NumberingCounter {
  constructor() {
    this.fig = 0
    this.tbl = 0
  }

  numberingFor(variable) {
    if (Label.isFigure(variable)) {
      return this.fig
    }
    else {
      return this.tbl;
    }
  }

  incrementTable() {
    ++this.tbl
  }

  incrementFigure() {
    ++this.fig
  }
}

function isContainer(element) {
  const containerTypes = new Set([
    DocumentApp.ElementType.BODY_SECTION,
    DocumentApp.ElementType.PARAGRAPH,
    DocumentApp.ElementType.TABLE,
    DocumentApp.ElementType.LIST_ITEM
  ])
  if (!element.getType) {
    return true
  }
  return containerTypes.has(element.getType())
}

class Visitor {
  constructor(callback) {
    this.callback = callback
  }

  visit(element) {
    const continueToVisit = this.callback(element)
    if (!continueToVisit) {
      return
    }
    if (!isContainer(element)) {
      return
    }
    this.visitChild(element)
  }

  visitChild(container) {
    const num = container.getNumChildren()
    for (var index = 0; index < num; ++index) {
      this.visit(container.getChild(index))
    }
  }
}

const Transition = {
  Fig: "Fig",
  Tbl: "Tbl",
  FigLabel: "FigLabel",
  TblLabel: "TblLabel",
  Chapter: "Chapter",
  NoTrans: "NoTrans",
  Start: "Start"
}

function toTransition(element) {
  const map = new Map([
    [DocumentApp.ElementType.INLINE_IMAGE, (element) => Transition.Fig],
    [DocumentApp.ElementType.TABLE, (element) => Transition.Tbl],
    [DocumentApp.ElementType.VARIABLE, (element) => {
      if (Label.isFigure(element)) { return Transition.FigLabel }
      if (Label.isTable(element)) { return Transition.TblLabel }
      return Transition.Start
    }],
    [DocumentApp.ElementType.PARAGRAPH, (element) => {
      if (element.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
        return Transition.Chapter
      }
      return Transition.NoTrans
    }],
    [DocumentApp.ElementType.BODY_SECTION, (element) => Transition.Start]
  ])

  const converter = map.get(element.getType())
  return converter ? converter(element) : Transition.NoTrans
}

class Context {
  constructor() {
    this.chapter = null
    this.numbering = new NumberingCounter()
  }
}

class State {
  constructor(name, expectedTransitions) {
    this.name = name
    this.expectedTransitions = expectedTransitions
  }

  expected(transition) {
    return !!this.expectedTransitions.find(expected => expected === transition)
  }

  enter(element, ctx) { }

  exit(transition, element, ctx) { }
}

class StFig extends State {
  constructor() {
    super(Transition.Fig, [Transition.FigLabel])
  }

  enter(element, ctx) {
    typeGuard(element, DocumentApp.ElementType.INLINE_IMAGE)
    ctx.numbering.incrementFigure()
  }

  exit(transition, element, ctx) {
    if (this.expected(transition)) {
      typeGuard(element, DocumentApp.ElementType.VARIABLE)
      Label.labeling(ctx.chapter, ctx.numbering, element)
    }
  }
}

class StTbl extends State {
  constructor() {
    super(Transition.Tbl, [])
  }
  enter(element, ctx) {
    typeGuard(element, DocumentApp.ElementType.TABLE)
  }
}

class StFigLabel extends State {
  constructor() {
    super(Transition.FigLabel, [])
  }
}

class StTblLabel extends State {
  constructor() {
    super(Transition.TblLabel, [Transition.Tbl])
    this.lastTableLabel = null
  }

  enter(element, ctx) {
    typeGuard(element, DocumentApp.ElementType.VARIABLE)
    ctx.lastTableLabel = element
  }

  exit(transition, element, ctx) {
    if (this.expected(transition)) {
      typeGuard(element, DocumentApp.ElementType.TABLE)
      ctx.numbering.incrementTable()
      Label.labeling(ctx.chapter, ctx.numbering, ctx.lastTableLabel)
    }
    this.lastTableLabel = null
  }
}

class StChapter extends State {
  constructor() {
    super(Transition.Chapter, [])
  }

  enter(element, ctx) {
    typeGuard(element, DocumentApp.ElementType.PARAGRAPH)
    if (element.getHeading() !== DocumentApp.ParagraphHeading.HEADING1) {
      throw new Error('Not a chapter')
    }
    ctx.chapter = element
    ctx.numbering = new NumberingCounter()
  }
}

class StateMachine {
  constructor() {
    const start = new State(Transition.Start, [])
    this.states = new Map([
      [Transition.Fig, new StFig()],
      [Transition.Tbl, new StTbl()],
      [Transition.FigLabel, new StFigLabel(Transition.FigLabel, [])],
      [Transition.TblLabel, new StTblLabel(Transition.TblLabel, [Transition.Tbl])],
      [Transition.Chapter, new StChapter(Transition.Chapter, [])],
      [Transition.Start, start]
    ])
    this.currentState = start
    this.ctx = new Context()
  }

  transition(element) {
    const transition = toTransition(element)
    if (transition === Transition.NoTrans) {
      return
    }
    this.currentState.exit(transition, element, this.ctx)
    this.currentState = this.states.get(transition) ?? this.currentState
    this.currentState.enter(element, this.ctx)
  }
}
