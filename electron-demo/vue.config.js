module.exports={
    pluginOptions:{
        electronBuilder:{
            nodeIntegration: true,
            preload:'src-electron/electron-preload.ts'
        }
    }
}