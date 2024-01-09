(async () => {
  const { sourceAuth } = require('./config')
  const { createSharepointClient } = require('./lib/sharepoint-client')
  const { writeFileSync, readFileSync } = require('fs')
  const { logger } = require('@vtfk/logger')
  const formatterArchive = require('./data/content-type-formatter-archive')
  const formatterPublish = require('./data/content-type-formatter-publish')
  const formatterNoPublish = require('./data/content-type-formatter-no-publish')

  const spClient = createSharepointClient({
    clientId: sourceAuth.clientId,
    pfxcert: readFileSync(sourceAuth.pfxPath).toString('base64'),
    thumbprint: sourceAuth.pfxThumbprint,
    tenantId: sourceAuth.tenantId,
    tenantName: sourceAuth.tenantName
  })

  /* Her legger man inn listeurl - samt hvilken formatter man skal bruke - resten skal gå av seg selv. */

  const listUrl = 'https://vtfk.sharepoint.com/sites/T-ORG-DIGI-Teknologiogutvikling/lists/ttumoter'
  const formatter = formatterPublish

  /* Det under trenger man ikke gjøre noe med */

  const { listId } = await spClient.getListAndSiteId(listUrl)

  const contentStyleResult = (await spClient.getListContentTypes(listUrl, listId, 'Element')).d.results // SP rest API returns "d" as main property, don't know why

  if (contentStyleResult.length !== 1) {
    throw new Error(`Found more than one contentStyle with name "Element" for list: ${listUrl}`)
  }
  const { ClientFormCustomFormatter, StringId } = contentStyleResult[0]

  if (JSON.stringify(formatter) === ClientFormCustomFormatter) {
    logger('info', [`Custom formatter was already correct for list ${listUrl}`])
    process.exit(0)
  }

  const updateCustomFormatterResult = await spClient.updateListContentType(listUrl, listId, StringId, formatter)
  writeFileSync('./ignore/updateCustomFormatterResult.json', JSON.stringify(updateCustomFormatterResult, null, 2))

})()
