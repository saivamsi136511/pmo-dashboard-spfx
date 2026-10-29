# PMO Dashboard

This SharePoint Framework (SPFx) 1.11 client-side web part displays project
data from the current SharePoint site. It reads the `PMO_Projects` list using
the current user's SharePoint permissions; it is not a separately hosted static
website.

## Build

This legacy SPFx version requires Node.js 10.x and TypeScript 3.3. The GitHub
Actions workflow uses Node 10.24.1 and installs the pinned dependencies from
`package-lock.json`.

```sh
npm ci
npm run build -- --ship
node_modules/.bin/gulp package-solution --ship
```

The packaged solution is written to
`sharepoint/solution/pmo-dashboard.sppkg`. The workflow uploads this file as
the `pmo-dashboard-sppkg` artifact after a successful build.

## Deploy to SharePoint

1. Push the changes to `main` and wait for **Build SPFx PMO Dashboard** to
   complete successfully.
2. In the successful workflow run, download the `pmo-dashboard-sppkg` artifact
   and extract `pmo-dashboard.sppkg`.
3. Upload the `.sppkg` file to your tenant or site SharePoint App Catalog.
4. In the App Catalog, select **Deploy**. If prompted, approve the deployment.
   The package includes its client-side assets.
5. On the SharePoint site where the dashboard should appear, create a list
   named `PMO_Projects` with the internal fields `Title`, `Project_Status`,
   and `Project_Manager` (Person or Group).
6. Add a modern SharePoint page, add the **PmoDashboard** web part, and publish
   the page. The current user's access to the site and list controls whether
   project data can be read.

`skipFeatureDeployment` is enabled, so a tenant administrator can make the
web part available across sites during deployment. Otherwise, add the app to
the target site as required by your App Catalog settings. The package is marked
as domain isolated in `config/package-solution.json`; confirm that the tenant
supports isolated web parts before deploying it.

## Local workbench

Install dependencies with Node 10.x, then run `gulp serve` and use the local
SharePoint workbench. A SharePoint tenant is needed to verify data access and
the required list schema.
