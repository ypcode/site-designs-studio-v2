## site-designs-studio-v2

The Site Design Studio v2 is a new version of the solution.
It is now a single app part page instead a WebPart.
It is designed to be used as an entire solution on a tenant to provision the customization and configuration to the sites allowing its users to create and manage Site Designs and Site Scripts without the need to write any Powershell nor JSON.

 > Still in beta release, a few other features are still under developement and not published yet

### MPA

- Build the package either in DEBUG or RELEASE config and add it to your app catalog
- From a modern site home page, click New > Page > App > Site Designs Studio

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

