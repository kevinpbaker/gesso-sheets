# Scripts: the security model

A script is code. A workbook can come from anyone, so a script can be
code a stranger wrote, and a spreadsheet that runs it is a spreadsheet
that runs a stranger's code on the reader's machine. This document says
what a script here is allowed to do, what stops it doing more, and what
is left undefended. Each answer has a spec in
`src/script/Scripts.spec.ts` that fails if the answer stops being true.

It was written before any script could be written, which is Part
Five's rule: the model comes first, and the feature is built inside it.

## What a script is

Source text and where it came from. A script is **typed** in this
sheet, or brought in **with a file**, and the file is named. The
origin travels with the script and is never inferred from the text.

`LAMBDA` is not a script. It is a formula, runs where formulas run,
and can do nothing a formula cannot. This document is about scripts.

## Where it runs

**In a worker of its own, one per run, thrown away after the run.**

- Not on the main thread, where the input is handled.
- Not in the render worker, whose frames are Part One's claim.
- Not in the application worker, which holds the workbook. A script
  reaches the workbook only through what it is handed.

The application worker starts it, as a worker nested inside itself,
so a run never crosses the main thread or the render worker: the
values go from the thread that has them to the worker that reads them,
and the changes come straight back. The host sends the worker one run
and terminates it when the run ends, however it ends. A second message
to the worker is ignored, and whatever the script left pending goes
with the worker. One run at a time: a second while one is going is
refused, and opening another workbook ends the one in progress.

*Spec:* a run starts one worker and ends it.

## What it can reach

**The `sheet` that was showing, the `workbook`, and a `console` that
writes to the run's log. Nothing else.** `src/script/api.ts` is the
whole of it, as types:

- `sheet.range('A1:C9')`, with `.values` as rows, `.write(rows)` or
  `.write(value)`, `.format({ … })` and `.clear()`; and
  `sheet.read('B4')` and `sheet.write('B4', value)` for one cell.
- `workbook.sheets`, `workbook.sheet(name)` and
  `workbook.addSheet(name)`.
- `console.log(…)`, which keeps 200 lines of 1,000 characters.

A value is text, a number, true or false, or nothing. A string is
written as if typed, so `'=A1*2'` becomes a formula. A format is one
of a listed few fields: bold, italic, underline, wrap, a colour or a
fill written as `#rrggbb`, an alignment, a font size from 6 to 72, and
a number format from a list of six. The run starts from a copy of
every sheet's values, so a script can read the whole workbook. It
cannot read what is not a value: formulas, formats, notes, names,
charts, the file's name, or anything outside the workbook.

The worker's first act, before a script's code exists, is SES's
`lockdown()`, from `ses` (Hardened JavaScript, from Agoric). It freezes
every shared intrinsic, so no script can change `Object.prototype`,
`Array`, or anything else the rest of the worker relies on. The script
is then evaluated in a `Compartment` whose only globals are the three it
is given.

That rules out, by construction and not by a list:

- the network: `fetch`, `XMLHttpRequest`, `WebSocket`, `EventSource`;
- storage: `indexedDB`, `caches`, `localStorage` (which a worker does
  not have anyway);
- the worker itself: `self`, `postMessage`, `onmessage`,
  `importScripts`, `navigator`, `location`;
- timers, the clock (`Date.now()` throws), and randomness
  (`Math.random()` throws);
- loading code. SES rejects `import(…)` and direct `eval` in the
  source, including through the `Function` constructor.

The last one is why this is SES and not a worker with its globals
deleted. Dynamic `import()` is syntax, not a global. It cannot be
deleted, and a worker that only removed `fetch` could still load and
run code from any address. A list of deleted globals is also a list
that the next browser feature silently lengthens. A compartment starts
from nothing.

*Specs:* the network (`fetch`, `WebSocket`), the worker (`self`,
`postMessage`), Node's reach (`process`, `require`), timers, the clock,
randomness, `import(…)`, `eval`, the `Function` escape, and tampering
with a prototype or with `sheet` itself — run in the real worker on a
Node thread, where each of them really exists. Checked in Chrome
twice. In Phase 29, the worker loaded on its own found `fetch`, `self`,
`postMessage`, `importScripts`, `indexedDB` and `navigator` all
`undefined`, rejected `import(…)` and the `Function` escape, and had an
infinite loop terminated at its time limit. In Phase 30, a script typed
into the editor and run as the application runs it, nested in the
application worker, found `fetch`, `XMLHttpRequest`, `WebSocket`,
`importScripts`, `postMessage`, `self`, `indexedDB`, `caches`,
`navigator`, `location`, `setTimeout` and `Worker` all `undefined`.

## How much it can do

- **Time.** A run has a limit, five seconds by default. At the limit
  the host terminates the worker, and none of the run's writes are
  kept: a run is all or nothing.
- **Stop.** The editor's Run becomes Stop while a run is going, and
  the status bar says a run is going while the editor is closed.
  Stopping ends the worker and keeps nothing.
- **Writes.** A run may write or format a bounded number of cells,
  100,000 counted together, and add at most 20 sheets. The worker
  stops the script at the limit, and the host counts again, because
  the worker is where the untrusted code is. The host believes nothing
  a worker reports. A change to a sheet that does not exist, a cell
  outside the sheet, a value no cell can hold, a format field or value
  the list above does not have, too much of anything, or a message of
  the wrong shape refuses the whole run rather than trimming it. The
  specs include workers that lie.
- **Undo.** A whole run is one step of undo, applied as one
  transaction on the application worker, so the grid redraws once at
  the end and not once a cell. Each write goes through the same path
  as typing, so a rule on a cell can refuse it, and the run says how
  many were. The one exception is a sheet the run added: adding a
  sheet by hand cannot be undone either, so undo takes back what the
  run wrote and the sheet stays, and the run says so.

*Specs:* the time limit, Stop, the limits in the worker, a host that
refuses each kind of lie, and a run of the real worker through the
sheet that fills a column, formats it and is undone in one step.

## Where scripts come from

- A script **typed** here runs when its author runs it.
- A script that came **with a file** is shown as the file's, not the
  person's, and runs only when somebody chooses to run it — each time,
  having been told where it came from. Opening the file never runs it,
  a recalculation never runs it, and neither does anything else. The
  host refuses such a run unless it is confirmed, without starting a
  worker. The editor's Run asks the question, and asks it again next
  time.
- **Every script in an opened file is the file's, whatever the file
  says.** A `.gsheet` records where each script came from, for the
  library's own copy, which only this sheet writes. When a file is
  opened that record is overwritten: a file that could say "typed
  here" would be a file that could skip the question. That includes a
  file this sheet saved, because a file on disk is one anybody could
  have edited since.
- A file's script stays the file's when it is edited. Otherwise
  changing one character would make a stranger's code the person's.
- Scripts are saved in `.gsheet` and not in `.xlsx`, because an Excel
  file's code is VBA and this is not.

*Specs:* a file's script is refused unconfirmed, starts no worker, and
runs when confirmed; an opened file's scripts are the file's even when
the file says otherwise; and the editor asks each time.

## What is left undefended

Said here rather than discovered later.

- **Memory.** A browser gives no way to limit a worker's memory. A
  script that allocates until the tab runs out can crash the tab. It
  cannot read anything or send anything while it does, so this is a
  script denying service to the person who chose to run it. The time
  limit bounds how long it can try.
- **Timing side channels.** A script has no clock, which removes the
  easy one. Another thread could be used as a clock, but a script
  cannot start one.
- **What the person asks for.** A script that someone chose to run can
  do anything to the workbook that the API allows: overwrite every
  cell it may write, within the limits. The model keeps a script from
  reaching past the workbook, not from being a bad script. Undo is the
  answer to a bad script.
- **SES itself.** The compartment is the boundary, and it is someone
  else's code. It is pinned to one version, and a bump is a change to
  this document's claims.

## Adding to it

Each new capability is an entry in `api.ts`, in the list of what a
script can reach above, in the host's checks, and in a spec. A
capability that reaches past the workbook is a change to this document
first.

---

# Functions in cells

`=TAX(B2, 0.2)`, where `TAX` is JavaScript written in the workbook:
Part Six. Everything above is about scripts, which somebody runs.
This is about functions, which a formula calls whenever its inputs
change, with nobody choosing each time. That difference is why
functions get a sandbox of their own and a narrower reach.
`src/script/CellFunctions.ts` is the interpreter's side, and
`CellFunctions.spec.ts` has a spec per answer.

## Where it runs

**In QuickJS, compiled to WebAssembly, on the application worker.**
The recalculation is synchronous and needs each answer before the next
cell, so a worker, which can only answer later, will not do. The
interpreter runs inside the thread that recalculates, and a call is an
ordinary synchronous call into it.

It is still a separate world. The interpreter has its own heap and its
own global object, and the host's objects do not exist in it. A
workbook's functions are one module in one context. Defining them again
makes a fresh context, so nothing the last version left in the global
object survives.

## What it can reach

**Its arguments, and nothing else.** A function is given values and
returns a value, and both cross as plain data: numbers, strings, true
or false, nothing, or rows of them. It reads no cell it was not handed.
It writes nothing. It cannot see the workbook, the sheet, or the cell
it is in.

Checked in the real interpreter, from inside a function. In Node,
`fetch`, `XMLHttpRequest`, `WebSocket`, `self`, `postMessage`,
`importScripts`, `indexedDB`, `Worker`, `performance`, `console`,
timers, `process`, `require`, and QuickJS's own `std` and `os` modules
are all `undefined`. In a Chrome worker, `fetch`, `self`,
`postMessage`, `importScripts`, `indexedDB`, `navigator`, `location`,
`Worker`, `WebSocket`, `XMLHttpRequest` and `caches` are too. `import(…)`
has no loader to answer it and fails. `eval` is the interpreter's own,
and reaches nothing new.

**Nothing that changes on its own.** A dependency graph trusts that a
function given the same arguments gives the same answer. So there is no
`Math.random`. `Date` makes a date from what it is told, and refuses to
say what the time is now: `Date.now()` and `new Date()` throw.

## How much it can do

The interpreter enforces three limits, and each is checked by a spec
that exceeds it and then makes another call, which works:

- **Time.** 50 ms per call, ended by the interpreter's interrupt
  handler. Measured in Chrome: an endless loop ended at 50.4 ms.
  Defining the module runs under the same deadline, so a module that
  loops at the top level is refused rather than hanging the thread.
- **Memory.** 32 MB for the workbook's functions together, which a
  browser worker cannot give a script. An allocation past it is an
  error in the call. The gap this document leaves open for scripts is
  closed for functions.
- **Stack.** 128 kB, about 740 calls deep. The limit is set below what
  the host thread survives. At 256 kB, endless recursion overflowed the
  host's own stack in Chrome and in Node, before the interpreter's
  check fired, which is an error on the workbook's thread rather than
  in a cell. 224 kB was the largest that held in Chrome with nothing
  else on the stack, and a call from the evaluator arrives with the
  evaluator's frames under it.

A result that is not a value (an object, a function) is an error, not a
value. Phase 32 turns each failure into an error value in the cell, and
adds a budget per recalculation slice as well as per call.

## Where functions come from

The same two places as scripts, and the same rule underneath: nothing
in a file runs by itself. A formula runs every time, so the rule
cannot be *ask each time*. It becomes this:

- Functions **written here** are on.
- Functions that came **with a file** are off until the person turns
  them on for that workbook. Off, a call is an error that says why.
- Turned on is remembered in this browser's library, **never in the
  file**. A file cannot say it has been trusted, including a file this
  sheet saved, because a file on disk is one anybody could have edited
  since.

What makes turning them on a smaller decision than running a script is
the reach above. A hostile function can spend its time and memory,
which are bounded, and it can return a wrong answer, which any formula
can do. It cannot write a cell, read one it was not given, or send
anything anywhere.

## What is left undefended

- **Wrong answers.** A function can compute anything from its
  arguments, including something misleading. That is true of every
  formula, and the defence is the same: the formula is there to read.
- **Time, within the limits.** Five thousand cells each spending most
  of 50 ms is minutes of recalculation. The budget per slice in Phase 32
  bounds how much of that the thread spends before it hands back, and
  the error values say which cells are slow.
- **QuickJS itself.** The interpreter and its WebAssembly build are the
  boundary, and they are someone else's code, `quickjs-emscripten`
  0.32.0. It is pinned, and a bump is a change to this section's
  claims.
