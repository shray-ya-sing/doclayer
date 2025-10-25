import { defineConfig } from 'vitepress'

export default defineConfig({
  title: "DocLayer",
  description: "Cross-platform PowerPoint generation library for C#, Python, and TypeScript",
  base: '/',
  
  themeConfig: {
    logo: '/logo.svg',
    
    nav: [
      { text: 'Home', link: '/' },
      { text: 'Guide', link: '/guide/getting-started' },
      { text: 'API Reference', link: '/api/csharp' },
      { text: 'GitHub', link: 'https://github.com/shray-ya-sing/doclayer' }
    ],

    sidebar: [
      {
        text: 'Getting Started',
        items: [
          { text: 'Introduction', link: '/guide/introduction' },
          { text: 'Installation', link: '/guide/installation' },
          { text: 'Quick Start', link: '/guide/getting-started' }
        ]
      },
      {
        text: 'API Reference',
        items: [
          { text: 'C# / .NET', link: '/api/csharp' },
          { text: 'Python', link: '/api/python' },
          { text: 'TypeScript', link: '/api/typescript' },
          { text: 'Web API', link: '/api/webapi' }
        ]
      },
      {
        text: 'Examples',
        items: [
          { text: 'Create Presentations', link: '/examples/create' },
          { text: 'Extract Content', link: '/examples/extract' },
          { text: 'Edit Slides', link: '/examples/edit' },
          { text: 'Render Images', link: '/examples/render' }
        ]
      },
      {
        text: 'Deployment',
        items: [
          { text: 'Docker', link: '/deployment/docker' },
          { text: 'Azure', link: '/deployment/azure' },
          { text: 'AWS', link: '/deployment/aws' }
        ]
      }
    ],

    socialLinks: [
      { icon: 'github', link: 'https://github.com/shray-ya-sing/doclayer' }
    ],

    footer: {
      message: 'Released under the MIT License.',
      copyright: 'Copyright © 2024-present DocLayer'
    },

    search: {
      provider: 'local'
    }
  }
})
