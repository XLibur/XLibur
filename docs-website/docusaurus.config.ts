import type * as Preset from '@docusaurus/preset-classic';
import type {Config} from '@docusaurus/types';
import {themes as prismThemes} from 'prism-react-renderer';

const organizationName = 'XLibur';
const projectName = 'XLibur';

const config: Config = {
  title: 'XLibur',
  tagline: 'Read, manipulate, and write Excel 2007+ files from .NET — no Excel required',
  favicon: 'img/icon.png',

  // Published to GitHub Pages at https://xlibur.github.io/XLibur/
  url: `https://${organizationName.toLowerCase()}.github.io`,
  baseUrl: `/${projectName}/`,
  organizationName,
  projectName,
  trailingSlash: false,

  onBrokenLinks: 'throw',
  markdown: {
    hooks: {
      onBrokenMarkdownLinks: 'warn',
    },
  },

  i18n: {
    defaultLocale: 'en',
    locales: ['en'],
  },

  presets: [
    [
      'classic',
      {
        docs: {
          // Docs-only mode: the docs are served from the site root.
          routeBasePath: '/',
          sidebarPath: './sidebars.ts',
          editUrl: `https://github.com/${organizationName}/${projectName}/tree/main/docs-website/`,
        },
        blog: false,
        theme: {
          customCss: './src/css/custom.css',
        },
      } satisfies Preset.Options,
    ],
  ],

  themes: [
    [
      '@easyops-cn/docusaurus-search-local',
      {
        // Offline search: the index is built at compile time and served as
        // static assets, so it works on GitHub Pages with no search backend.
        hashed: true,
        // The site runs in docs-only mode (docs.routeBasePath === '/'), so the
        // indexer has to be pointed at the root too — it defaults to '/docs'
        // and would otherwise index nothing.
        docsRouteBasePath: '/',
        indexBlog: false,
        highlightSearchTermsOnTargetPage: true,
      },
    ],
  ],

  themeConfig: {
    image: 'img/logo.png',
    colorMode: {
      respectPrefersColorScheme: true,
    },
    navbar: {
      title: 'XLibur',
      logo: {
        alt: 'XLibur logo',
        src: 'img/icon.png',
      },
      items: [
        {
          type: 'docSidebar',
          sidebarId: 'docsSidebar',
          position: 'left',
          label: 'Docs',
        },
        {
          href: 'https://www.nuget.org/packages/XLibur.Bundle',
          label: 'NuGet',
          position: 'right',
        },
        {
          href: `https://github.com/${organizationName}/${projectName}`,
          label: 'GitHub',
          position: 'right',
        },
      ],
    },
    footer: {
      style: 'dark',
      links: [
        {
          title: 'Docs',
          items: [
            {label: 'Introduction', to: '/'},
            {label: 'Getting Started', to: '/getting-started'},
          ],
        },
        {
          title: 'Packages',
          items: [
            {label: 'XLibur.Bundle', href: 'https://www.nuget.org/packages/XLibur.Bundle'},
            {label: 'XLibur', href: 'https://www.nuget.org/packages/XLibur'},
          ],
        },
        {
          title: 'More',
          items: [
            {label: 'GitHub', href: `https://github.com/${organizationName}/${projectName}`},
            {
              label: 'Issues',
              href: `https://github.com/${organizationName}/${projectName}/issues`,
            },
            {label: 'ClosedXML docs', href: 'https://closedxml.github.io/ClosedXML/'},
          ],
        },
      ],
      copyright: `Copyright © ${new Date().getFullYear()} XLibur contributors. MIT licensed.`,
    },
    prism: {
      theme: prismThemes.github,
      darkTheme: prismThemes.dracula,
      additionalLanguages: ['csharp', 'bash', 'powershell'],
    },
  } satisfies Preset.ThemeConfig,
};

export default config;
