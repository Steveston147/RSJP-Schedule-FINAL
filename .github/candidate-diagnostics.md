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
  1012:5  warning  Unused eslint-disable directive (no problems were reported from 'react-hooks/exhaustive-deps')

✖ 1 problem (0 errors, 1 warning)
  0 errors and 1 warning potentially fixable with the `--fix` option.

ESLint found too many warnings (maximum: 0).
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
✓ Running next.config.js took 19ms
⚠ No build cache found. Please configure build caching for faster rebuilds. Read more: https://nextjs.org/docs/messages/no-cache
Attention: Next.js now collects completely anonymous telemetry regarding usage.
This information is used to shape Next.js' roadmap and prioritize features.
You can learn more, including how to opt-out if you'd not like to participate in this anonymous program, by visiting the following URL:
https://nextjs.org/telemetry


  Creating an optimized production build ...
Browserslist: browsers data (caniuse-lite) is 7 months old. Please run:
  npx update-browserslist-db@latest
  Why you should do it regularly: https://github.com/browserslist/update-db#readme
✓ Compiled successfully in 5.3s
  Running TypeScript ...

  We detected TypeScript in your project and reconfigured your tsconfig.json file for you.
  The following suggested values were added to your tsconfig.json. These values can be changed to fit your project's needs:

  	- include was updated to add '.next/dev/types/**/*.ts'

  The following mandatory changes were made to your tsconfig.json:

  	- jsx was set to react-jsx (next.js uses the React automatic runtime)

  Finished TypeScript in 3.0s ...
  Collecting page data using 3 workers ...
  Generating static pages using 3 workers (0/9) ...
  Generating static pages using 3 workers (2/9) 
  Generating static pages using 3 workers (4/9) 
  Generating static pages using 3 workers (6/9) 
✓ Generating static pages using 3 workers (9/9) in 257ms
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

