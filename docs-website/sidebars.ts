import type {SidebarsConfig} from '@docusaurus/plugin-content-docs';

const sidebars: SidebarsConfig = {
  docsSidebar: [
    'introduction',
    'getting-started',
    {
      type: 'category',
      label: 'Working with workbooks',
      collapsed: false,
      items: [
        'worksheets',
        'cells-and-ranges',
        'importing-exporting',
      ],
    },
    {
      type: 'category',
      label: 'Formatting',
      collapsed: false,
      items: [
        'styling',
        'theming',
        'conditional-formatting',
        'fonts',
      ],
    },
    {
      type: 'category',
      label: 'Data features',
      collapsed: false,
      items: [
        'tables',
        'autofilter',
        'pivot-tables',
        'data-validation',
      ],
    },
    {
      type: 'category',
      label: 'Calculation',
      collapsed: false,
      items: [
        'formulas',
        'functions',
      ],
    },
    {
      type: 'category',
      label: 'Output',
      collapsed: false,
      items: [
        'page-setup',
      ],
    },
  ],
};

export default sidebars;
