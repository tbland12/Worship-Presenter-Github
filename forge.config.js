module.exports = {
  outDir: 'out-publish',
  packagerConfig: {
    asar: true
  },
  makers: [
    {
      name: '@electron-forge/maker-squirrel',
      config: {}
    },
    {
      name: '@electron-forge/maker-zip',
      platforms: ['win32']
    }
  ],
  publishers: [
    {
      name: '@electron-forge/publisher-github',
      config: {
        repository: {
          owner: 'tbland12',
          name: 'Worship-Presenter-Github'
        },
        draft: true,
        prerelease: false
      }
    }
  ]
};
