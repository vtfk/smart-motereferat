const { getSharepointToken } = require('./get-sharepoint-token')
const { getWebUrlParts } = require('./graph-actions')
const axios = require('./axios-instance')()
const { logger } = require('@vtfk/logger')

const getColumns = async (libraryUrl, listId) => {
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`
  const query = `_api/web/lists(guid'${listId}')/fields?$select=Id,CustomFormatter,InternalName,StaticName,Title`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const { data } = await axios.get(`${baseUrl}/${query}`, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' } })
  return data
}

const modifyColumn = async (libraryUrl, listId, columnId, body) => {
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`
  const query = `_api/web/lists(guid'${listId}')/fields('${columnId}')`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const { data } = await axios.post(`${baseUrl}/${query}`, body, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-HTTP-Method': 'MERGE' } })
  return data
}

const getViews = async (libraryUrl, listId) => {
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`
  const query = `_api/web/lists(guid'${listId}')/views?$expand=ViewFields`// ?$select=Id,CustomFormatter,InternalName,StaticName,Title`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const { data } = await axios.get(`${baseUrl}/${query}`, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' } })
  return data
}

const addView = async (libraryUrl, listId, viewTitle) => {
  // https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-rest-reference/dn531433%28v%3doffice.15%29#viewfieldcollection-methods
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`
  const query = `_api/web/lists(guid'${listId}')/views?$expand=ViewFields`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const viewBody = {
    __metadata: { type: 'SP.View' },
    Title: viewTitle,
    PersonalView: false
  }

  const { data } = await axios.post(`${baseUrl}/${query}`, viewBody, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' } })
  return data
}

const removeView = async (libraryUrl, listId, viewTitle) => {
  // https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-rest-reference/dn531433%28v%3doffice.15%29#viewfieldcollection-methods
  // NOTE THAT THIS function only removes the first occurence of the viewTitle - if you have several views with the same title, only one of them are removed (could fix if though...)
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`
  const query = `_api/web/lists(guid'${listId}')/views/getbytitle('${viewTitle}')`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const { data } = await axios.post(`${baseUrl}/${query}`, {}, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', "X-HTTP-Method": "DELETE" } })
  return data
}

const removeViewField = async (libraryUrl, listId, viewId, fieldName) => {
  // https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-rest-reference/dn531433%28v%3doffice.15%29#viewfieldcollection-methods
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`
  const query = `_api/web/lists(guid'${listId}')/views('${viewId}')/viewfields/removeviewfield('${fieldName}')`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const viewBody = {}

  const { data } = await axios.post(`${baseUrl}/${query}`, viewBody, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' } })
  return data
}

const addViewField = async (libraryUrl, listId, viewId, fieldName) => {
  // https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-rest-reference/dn531433%28v%3doffice.15%29#viewfieldcollection-methods
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`
  const query = `_api/web/lists(guid'${listId}')/views('${viewId}')/viewfields/addviewfield('${fieldName}')`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const viewBody = {}

  const { data } = await axios.post(`${baseUrl}/${query}`, viewBody, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' } })
  return data
} 

const upsertView = async (libraryUrl, listId, view) => {
  if (!view.title) throw new Error('missing required parameter "view.title"')
  if (!view.columns || !Array.isArray(view.columns)) throw new Error('missing required parameter "view.columns" of type Array')
  
  // Hent views for en liste
  logger('info', ['upsertView', `Fetching views for ${libraryUrl} -  list: ${listId}`])
  const views = (await getViews(libraryUrl, listId)).d.results
  logger('info', ['upsertView', `Got ${views.length} views for ${libraryUrl}`])

  // Sjekk om view finnes basert på title
  let viewData = views.find(v => v.Title === view.title)

  // Om den ikke finnes lag den
  if (!viewData) {
    logger('info', ['upsertView', `View ${view.title} does not exist in ${libraryUrl}. Creating`])
    viewData = (await addView(libraryUrl, listId, view.title)).d
    logger('info', ['upsertView', `View ${view.title} succesfully created in ${libraryUrl}`])
  } else {
    logger('info', ['upsertView', `View ${view.title} already exists in ${libraryUrl}. No need to create`])
  }

  const viewColumns = viewData.ViewFields.Items.results
  const columnsToAdd = view.columns.filter(col => !viewColumns.includes(col))
  logger('info', ['upsertView', columnsToAdd.length > 0 ? `Need to add ${columnsToAdd.length} columns to view "${view.title}" in ${libraryUrl}` : `All required columns already exist in view "${view.title}" in ${libraryUrl}`])

  for (const column of columnsToAdd) {
    logger('info', ['upsertView', `adding column "${column} to view "${view.title} in ${libraryUrl}"`])
    await addViewField(libraryUrl, listId, viewData.Id, column)
    logger('info', ['upsertView', `sucessfylly added column "${column}" to view "${view.title} in ${libraryUrl}"`])
  }

  const columnsToRemove = viewColumns.filter(col => !view.columns.includes(col))
  logger('info', ['upsertView', columnsToRemove.length > 0 ? `Need to remove ${columnsToRemove.length} columns to view "${view.title}" in ${libraryUrl}` : `No columns need to be removed from view "${view.title}" in ${libraryUrl}`])

  for (const column of columnsToRemove) {
    logger('info', ['upsertView', `removing column "${column}" from view "${view.title} in ${libraryUrl}"`])
    await removeViewField(libraryUrl, listId, viewData.Id, column)
    logger('info', ['upsertView', `sucessfylly removed "${column} from view "${view.title} in ${libraryUrl}"`])
  }

  return 'Yes man!'
}

const cleanUpDefaultView = async (libraryUrl, listId, removeFields) => {
  // Get default view
  logger('info', ['cleanUpDefaultView', `Fetching views for ${libraryUrl} -  list: ${listId}`])
  const views = (await getViews(libraryUrl, listId)).d.results
  const defaultView = views.find(v => v.DefaultView)
  if (!defaultView) throw new Error(`Could not find default view for ${libraryUrl} in list ${listId}`)
  logger('info', ['cleanUpDefaultView', `Got default view: ${defaultView.Title} for ${libraryUrl}`])
  
  // Remove viewfields specified by function
  for (const removeField of removeFields) {
    logger('info', ['upsertView', `removing column "${removeField}" from view "${defaultView.title} in ${libraryUrl}"`])
    await removeViewField(libraryUrl, listId, defaultView.Id, removeField)
    logger('info', ['upsertView', `sucessfylly removed "${removeField} from view "${defaultView.title} in ${libraryUrl}"`])
  }

  return 'Yes man!'

}

const addComment = async (libraryUrl, listId, elementId) => {
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`

  const payload = {
    "__metadata": {
      "type": "Microsoft.SharePoint.Comments.comment"
    },
    "text": "Sjekk ut detta da @mention{0}. Det er kult",
    "mentions": {
      "__metadata": {
        "type" : "Collection(Microsoft.SharePoint.Comments.Client.Identity)"
      },
      results: [{email: 'epost@domene.no'}]
    }
  }

  const query = `_api/web/lists('${listId}')/items('${elementId}')/Comments()`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const { data } = await axios.post(`${baseUrl}/${query}`, payload, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' } })
  return data
}

const getListContentTypes = async (libraryUrl, listId, contentTypeName) => {
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`

  let query = `_api/web/lists('${listId}')/contenttypes`
  if (contentTypeName) query += `?$filter=Name eq '${contentTypeName}'`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const { data } = await axios.get(`${baseUrl}/${query}`, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' } })
  return data
}

/**
 * 
 * @param {*} libraryUrl 
 * @param {*} listId 
 * @param {*} contentTypeId 
 * @param {Object} ClientFormCustomFormatter { headerJSONFormatter: {hei: "hade"}, footerJSONFormatter: "", bodyJSONFormatter: "" }
 * @returns result
 */
const updateListContentType = async (libraryUrl, listId, contentTypeId, ClientFormCustomFormatter) => {
  const { siteName, tenantName } = getWebUrlParts(libraryUrl)
  const sharepointToken = await getSharepointToken()
  const baseUrl = `https://${tenantName}.sharepoint.com/sites/${siteName}`

  const query = `_api/web/lists('${listId}')/contenttypes('${contentTypeId}')`
  logger('info', ['Calling Sharepoint rest api', 'resource', `${baseUrl}/${query}`])

  const payload = {
    __metadata: {
      type: "SP.ContentType"
    },
    ClientFormCustomFormatter: JSON.stringify(ClientFormCustomFormatter)
  }

  const { data } = await axios.post(`${baseUrl}/${query}`, payload, { headers: { Authorization: `Bearer ${sharepointToken}`, Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-HTTP-Method': 'MERGE' } })
  return data
}


// When adding columns - we should have option not to add it automatically to defualt view
// Maybe just have function for cleanupDefault view 

module.exports = { modifyColumn, getColumns, getViews, addView, removeView, removeViewField, addViewField, upsertView, cleanUpDefaultView, addComment, getListContentTypes, updateListContentType }
