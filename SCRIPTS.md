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

The host starts the worker, sends it one run, and terminates it when
the run ends however it ends. A second message to the worker is
ignored, and whatever the script left pending goes with the worker.

*Spec:* a run starts one worker and ends it.

## What it can reach

**The `sheet` it is given, and a `console` that writes to the run's
log. Nothing else.**

The worker's first act, before a script's code exists, is SES's
`lockdown()`, from `ses` (Hardened JavaScript, from Agoric). It freezes
every shared intrinsic, so no script can change `Object.prototype`,
`Array`, or anything else the rest of the worker relies on. The script
is then evaluated in a `Compartment` whose only globals are the two it
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
Node thread, where each of them really exists. Checked in Chrome: the
same worker, loaded as the app will load it, finds `fetch`, `self`,
`postMessage`, `importScripts`, `indexedDB` and `navigator` all
`undefined`, rejects `import(…)` and the `Function` escape, and has an
infinite loop terminated at its time limit.

## How much it can do

- **Time.** A run has a limit, five seconds by default. At the limit
  the host terminates the worker, and none of the run's writes are
  kept: a run is all or nothing.
- **Writes.** A run may write a bounded number of cells. The worker
  stops the script at the limit, and the host counts again, because
  the worker is where the untrusted code is. The host believes nothing
  a worker reports. A write outside the sheet, a write of a value no
  cell can hold, too many writes, or a message of the wrong shape
  refuses the whole run rather than trimming it. The specs include
  workers that lie.
- **Undo.** A whole run is one step of undo. The host hands back the
  writes it accepts as one list, and Phase 30 applies that list as one
  transaction on the application worker.

*Specs:* the time limit, the write limit in the worker, and a host
that refuses each kind of lie.

## Where scripts come from

- A script **typed** here runs when its author runs it.
- A script that came **with a file** is shown as the file's, not the
  person's, and runs only when somebody chooses to run it — each time,
  having been told where it came from. Opening the file never runs it,
  a recalculation never runs it, and neither does anything else. The
  host refuses such a run unless it is confirmed, without starting a
  worker.

*Spec:* a file's script is refused unconfirmed, starts no worker, and
runs when confirmed.

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

## For the phase that builds on this

Phase 30 adds the API — ranges, formats, sheets — and the editor. Each
new capability is an entry in the list of what a script can reach, and
a spec. A capability that reaches past the workbook is a change to
this document first.
