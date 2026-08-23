# Next 16 candidate diagnostics

## audit
```text
found 0 vulnerabilities
```

## security
```text

> schedule@0.1.0 validate:security
> node scripts/validate-security-boundary.mjs

Security boundary validation passed.
```

## lint
```text

> schedule@0.1.0 lint
> eslint . --max-warnings=0


/home/runner/work/RSJP-Schedule-FINAL/RSJP-Schedule-FINAL/components/ScheduleApp.tsx
   998:6  warning  React Hook useMemo has a missing dependency: 'selectedProgram'. Either include it or remove the dependency array                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             react-hooks/exhaustive-deps
  1003:5  error    Error: Calling setState synchronously within an effect can trigger cascading renders

Effects are intended to synchronize state between React and external systems such as manually updating the DOM, state management libraries, or other platform APIs. In general, the body of an effect should do one or both of the following:
* Update external systems with the latest state from React.
* Subscribe for updates from some external system, calling setState in a callback function when external state changes.

Calling setState synchronously within an effect body causes cascading renders that can hurt performance, and is not recommended. (https://react.dev/learn/you-might-not-need-an-effect).

/home/runner/work/RSJP-Schedule-FINAL/RSJP-Schedule-FINAL/components/ScheduleApp.tsx:1003:5
  1001 |   useEffect(() => {
  1002 |     if (!selectedProgram) return;
> 1003 |     setStudentsCount(selectedProgram.studentsCount);
       |     ^^^^^^^^^^^^^^^^ Avoid calling setState() directly within an effect
  1004 |     setNewDate(selectedProgram.startDate);
  1005 |   }, [selectedProgram?.id]);
  1006 |                                                                     react-hooks/set-state-in-effect
  1005:6  warning  React Hook useEffect has a missing dependency: 'selectedProgram'. Either include it or remove the dependency array                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           react-hooks/exhaustive-deps
  1010:5  error    Error: Calling setState synchronously within an effect can trigger cascading renders

Effects are intended to synchronize state between React and external systems such as manually updating the DOM, state management libraries, or other platform APIs. In general, the body of an effect should do one or both of the following:
* Update external systems with the latest state from React.
* Subscribe for updates from some external system, calling setState in a callback function when external state changes.

Calling setState synchronously within an effect body causes cascading renders that can hurt performance, and is not recommended. (https://react.dev/learn/you-might-not-need-an-effect).

/home/runner/work/RSJP-Schedule-FINAL/RSJP-Schedule-FINAL/components/ScheduleApp.tsx:1010:5
  1008 |   useEffect(() => {
  1009 |     if (!selectedProgram) return;
> 1010 |     setOverrideDate(selectedProgram.startDate);
       |     ^^^^^^^^^^^^^^^ Avoid calling setState() directly within an effect
  1011 |     loadOverrideFromProgram(selectedProgram.startDate);
  1012 |     // eslint-disable-next-line react-hooks/exhaustive-deps
  1013 |   }, [selectedProgram?.id]);  react-hooks/set-state-in-effect
  1349:6  warning  React Hook useMemo has a missing dependency: 'selectedProgram'. Either include it or remove the dependency array                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             react-hooks/exhaustive-deps

✖ 5 problems (2 errors, 3 warnings)

```

## typecheck
```text

> schedule@0.1.0 typecheck
> tsc --noEmit

```

## build
```text

> schedule@0.1.0 build
> next build

▲ Next.js 16.3.2 (Turbopack)
✓ Running next.config.js took 18ms
⚠ No build cache found. Please configure build caching for faster rebuilds. Read more: https://nextjs.org/docs/messages/no-cache
Attention: Next.js now collects completely anonymous telemetry regarding usage.
This information is used to shape Next.js' roadmap and prioritize features.
You can learn more, including how to opt-out if you'd not like to participate in this anonymous program, by visiting the following URL:
https://nextjs.org/telemetry


  Creating an optimized production build ...
Browserslist: browsers data (caniuse-lite) is 7 months old. Please run:
  npx update-browserslist-db@latest
  Why you should do it regularly: https://github.com/browserslist/update-db#readme
✓ Compiled successfully in 5.1s
  Running TypeScript ...

  We detected TypeScript in your project and reconfigured your tsconfig.json file for you.
  The following suggested values were added to your tsconfig.json. These values can be changed to fit your project's needs:

  	- include was updated to add '.next/dev/types/**/*.ts'

  The following mandatory changes were made to your tsconfig.json:

  	- jsx was set to react-jsx (next.js uses the React automatic runtime)

  Finished TypeScript in 2.9s ...
  Collecting page data using 3 workers ...
  Generating static pages using 3 workers (0/9) ...
  Generating static pages using 3 workers (2/9) 
  Generating static pages using 3 workers (4/9) 
  Generating static pages using 3 workers (6/9) 
✓ Generating static pages using 3 workers (9/9) in 280ms
  Finalizing page optimization ...

Route (app)
┌ ○ /
├ ○ /_not-found
├ ƒ /api/auth/config
├ ƒ /api/auth/login
├ ƒ /api/auth/logout
├ ƒ /api/auth/session
└ ○ /login


ƒ Proxy (Middleware)

○  (Static)   prerendered as static content
ƒ  (Dynamic)  server-rendered on demand

```

