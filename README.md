# gessosheet

A Gesso application. The interface is built, laid out, painted and
hit-tested in a render worker; the page's own thread creates the canvas,
forwards input and does nothing else.

```bash
pnpm install    # or npm install
pnpm dev
```

Then `pnpm build` for a production bundle, `pnpm preview` to serve it,
and `pnpm typecheck` to check the types without building.

## The two pages

`/` is the spreadsheet. That is the whole of it: a grid, a menu bar, a
toolbar and a formula bar, and nothing on the screen that is about the
screen.

`/proof` is the same spreadsheet with the instruments attached — the
black strip along the top, and the one command the strip exists to
measure:

- a pulse driven by the main thread's own `requestAnimationFrame`, so
  it is alive exactly when that thread is;
- a button that holds the main thread in a busy loop for five seconds,
  which freezes the page and not the sheet;
- the layout heatmap, which washes every node the engine measured this
  frame and explains whatever is under the pointer;
- the render worker's frame rate, its worst gap, and how many nodes
  were re-measured;
- **Recalculate 200,000 cells** — on the toolbar, in the Data menu and
  on F9, and on this page only.

That list is why there are two urls. Nobody opening a spreadsheet
wants a black bar of instrumentation across the top of it, and a menu
item that exists to be photographed belongs on the page that
photographs it. `pnpm proof` drives `/proof`, because everything it
reads is there.

The routes are declared in `src/app/Routes.tsx`, in the render worker,
because a route names a component and a component cannot cross a
`postMessage`; the shell's half is `history: { mode: 'path' }` in
`createApp` and nothing else. The strip is the one part that cannot
work that way — it is DOM on the main thread, which is the entire
reason it is believable — so the shell reads the url for itself too,
through `src/route.ts`, the one module both threads share.

## The three files

| File            | What it is                                               |
| --------------- | -------------------------------------------------------- |
| `src/main.ts`   | The main thread: create the app and mount it into `#app` |
| `src/worker.ts` | The render worker: name the root component               |
| `src/App.tsx`   | The screen                                               |

There are three rather than two because a component cannot cross
`postMessage`, so the root has to be named on the worker's side of the
barrier.

`main.ts` names no worker, and that is `gesso-vite-plugin` in
`vite.config.ts`. It finds `worker.ts` beside `main.ts` and writes the
construction, which has to be written out as a literal because a bundler
emits a chunk for a worker it can see constructed and cannot see through
a variable holding the URL. The plugin also, while the dev server is
running:

- replaces the screen when you save `App.tsx`, instead of reloading the
  page, and puts the keyboard focus back where it was;
- draws what the render worker threw over the application that was
  running when it threw it, source-mapped, with the node named as a path
  through your components rather than as an id;
- says so when a save is about to reload the page, which happens when a
  module is reachable from the main thread as well as from the worker.

If you would rather say it out loud, write it and the plugin leaves the
construction alone:

```ts
createApp({
  renderWorker: () => new Worker(new URL('./worker.ts', import.meta.url), { type: 'module' })
});
```

```json
"jsx": "react-jsx",
"jsxImportSource": "gesso-framework"
```

Those two lines in `tsconfig.json` are the whole of what makes `<row>`
and `<text>` work. JSX here is a spelling rather than a runtime: it
compiles onto the element factories and produces the identical tree, so
`Row({ gap: 8 }, Text({ text: 'Ready' }))` is the same thing written the
other way.

## Icons

[Heroicons](https://heroicons.com/) is this project's icon set — the
24×24 outline half of it. It is MIT licensed, and the licence travels
with the path data in `src/app/heroicons.ts`.

Nothing in that package is imported at runtime. There is no DOM under
this interface, so an `<svg>` has nowhere to go; what Gesso's `Icon`
draws is the path itself, and `pnpm icons` lifts the paths this
application uses out of the package into `src/app/heroicons.ts`:

```bash
pnpm add -D heroicons@latest && pnpm icons
```

`scripts/icons.ts` holds the list of glyphs — add a line there and run
it again. `src/app/icons.ts` is the layer above, where a glyph gets an
application's name for it (`undo`, not `arrowUturnLeft`), and it is the
file to read to see which picture each button is wearing. A spec checks
the lifted data against the installed package, so an upgrade that skips
`pnpm icons` fails the build rather than quietly drawing last year's
glyph.

## Reporting errors from a production build

The overlay is a development tool and a production build carries no
reference to it. What a shipped application wants instead is its own
`onError`, which takes the same three arguments:

```ts
createApp({
  onError: (message, stack, source) => {
    reportToYourService({ message, stack, source });
  }
});
```

`source` is one of `message`, `uncaught`, `renderer`, `listener` or
`channel`, and it says what the failure cost: a `listener` error means
one handler did not run, a `renderer` error means the surface stopped
being updated, and a `channel` error means the data behind an intact
view has stopped arriving.

## Where to go next

- `App.tsx` is commented with what each part of it is doing.
- `gesso-components` has the controls: inputs, overlays, structure,
  data and media. `Switch` in `App.tsx` is one of them.
- Every prop takes a value or an Observable of that value. That is the
  whole binding model, and it is why the component body runs once.

## Why `vendor/` exists, and how to remove it

This project was scaffolded with `--local`, so it installs Gesso from
a checkout rather than from the registry: the packages were packed
into `vendor/` and the manifest points at the tarballs.

The packed packages declare each other by version range, so without
help a package manager is free to go looking for `gesso-core@^0.1.0`
on the registry and get a different copy than the one beside it.
`overrides` in `package.json` is what tells npm; `overrides` in
`pnpm-workspace.yaml` is what tells pnpm, which reads it nowhere else.

To pick up a further change, run `create-gesso-app --local` over this
directory again with `--force`.

To move to the published packages: delete `vendor/`, delete
`pnpm-workspace.yaml`, delete `overrides`, and put version ranges back
in `dependencies`.
