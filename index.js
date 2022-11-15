// Libreria para interactuar con excel
const ExcelJS = require('exceljs');

// Libreria para hacer requests HTTP
const needle = require('needle');

needle.defaults({
  open_timeout: 30000,
});

// variables globales
const organization = 'master-aws';
const apiVersion = '7.0';
const personalAccessToken = 'xxx';

const fechaInicio = "2022-10-01";
const fechaFin = "2022-10-31";

/**
 * En esta constante se especifican los headers de la tabla
 */
const columns = [
  { name: 'Release ID', filterButton: true },
  { name: 'Release Name', filterButton: true },
  { name: 'Artifact Name/Alias', filterButton: true },
  { name: 'Artifact Definition Name', filterButton: true },
  { name: 'Repository Name', filterButton: true },
  { name: 'Requested For Display Name', filterButton: true },
  { name: 'Requested For Unique Name', filterButton: true },
  { name: 'Release Definition Name', filterButton: true },
  { name: 'Release Environment Name', filterButton: true },
  { name: 'Pre deploy Approvals Approver Display Names', filterButton: true },
  { name: 'Pre deploy Approvals Approved by Display Names', filterButton: true },
  { name: 'Pre deploy Approvals Approved by Unique Names', filterButton: true },
  { name: 'Pre deploy Approvals Modified Date', filterButton: false },
]

/**
 * 
 * @param {string} accessToken token que se quiere codificar a base64 
 * @returns token codificado a base64
 */
const encodeToken = (accessToken) => {
  const buffer = Buffer.from(`:${accessToken}`)
  const token = buffer.toString('base64');
  return token;
}

/**
 * 
 * @param {string} accessToken El Personal Access Token sacado de Azure DevOps 
 */
const getProjects = async (accessToken) => {
  let projects = [];
  const apiURL = `https://dev.azure.com/${organization}/_apis/projects?api-version=${apiVersion}&queryOrder=ascending`;
  const encodedToken = encodeToken(accessToken);

  const options = {
    headers: { 'Authorization': `Basic ${encodedToken}` }
  }

  try {
    let response = await needle('get', apiURL, null, options)

    if (response.statusCode >= 200 && response.statusCode < 300) {
      response.body.value.forEach((value) => {
        projects.push({
          projectId: value.id,
          projectName: value.name
        })
      })
      return projects;
    } else {
      console.log("Ocurrio un error en el request")
      console.log("statusCode", response.statusCode)
      console.log("statusMessage", response.statusMessage)
      return null
    }

  } catch (error) {
    console.log("ocurrio un error", error)
    return error
  }

}

const getDeployments = async (projectId, accessToken) => {
  let deployments = [];
  const apiURL = `https://vsrm.dev.azure.com/${organization}/${projectId}/_apis/release/deployments?` +
    `api-version=${apiVersion}&` +
    `deploymentStatus=succeeded&` +
    `$top=200&` +
    `minStartedTime=${fechaInicio}&` +
    `maxStartedTime=${fechaFin}&` +
    `queryOrder=ascending`;

  console.log(apiURL)

  const encodedToken = encodeToken(accessToken);

  const options = {
    headers: { 'Authorization': `Basic ${encodedToken}` }
  }

  try {
    let response = await needle('get', apiURL, null, options)
    if (response.statusCode >= 200 && response.statusCode < 300) {
      deployments = response.body.value
      return deployments;
    } else {
      console.log("Ocurrio un error en el request")
      console.log(console.log("statusCode", response.statusCode))
      console.log(console.log("statusMessage", response.statusMessage))
      return null
    }

  } catch (error) {
    console.log("ocurrio un error", error)
    if (error.code == 'ECONNRESET') console.log("ECONNRESET")
    throw error
  }

}



/**
 * 
 * @param {*} item Este Item representa un objeto del atributo 'value' del request API deployments
 * @returns Retorna un array de atributos a ser usados para representar una fila de excel, con los datos del deployment
 */
const getRowFromItem = (item) => {
  // crear un array vacio e ir poblandolo poco a poco
  let items = [];
  // Si el item esta vacio, enviarlo sin datos
  if (!item) return items;

  // Inicializar 3 columnas del item en vacio, ya que algunos items no tienen ningun predeploy approvals
  let preDeployApprovalsApproverDisplayName = "";
  let preDeployApprovalsApprovedByDisplayName = "";
  let preDeployApprovalsApprovedByUniqueName = "";
  let preDeployApprovalsModifiedDate = "";

  // Si no hay predeploy approvals, entonces saltar este paso
  if (item.preDeployApprovals) {

    // Iterar entre los predeploy approvals que tiene el despliegue
    item.preDeployApprovals.forEach((approval, index) => {

      // Si tiene approver, entonces llenar ese dato en el string
      if (approval.approver) {
        // Si es el primer item, entonces asignar el valor
        if (preDeployApprovalsApproverDisplayName.length === 0) {
          if (approval.approver) preDeployApprovalsApproverDisplayName = approval.approver.displayName;
          // Si no es el primer item, concatenar el string anterior con el nuevo, con una "," en medio
        } else {
          if (approval.approver) preDeployApprovalsApproverDisplayName = preDeployApprovalsApproverDisplayName.concat(', ', approval.approver.displayName)
        }
        // Si es el primer item, entonces asignar el valor
        if (preDeployApprovalsApprovedByDisplayName.length === 0) {
          if (approval.approvedBy) preDeployApprovalsApprovedByDisplayName = approval.approvedBy.displayName;
          // Si no es el primer item, concatenar el string anterior con el nuevo, con una "," en medio
        } else {
          if (approval.approvedBy) preDeployApprovalsApprovedByDisplayName = preDeployApprovalsApprovedByDisplayName.concat(', ', approval.approvedBy.displayName)
        }
        // Si es el primer item, entonces asignar el valor
        if (preDeployApprovalsApprovedByUniqueName.length === 0) {
          if (approval.approvedBy) preDeployApprovalsApprovedByUniqueName = approval.approvedBy.uniqueName;
          // Si no es el primer item, concatenar el string anterior con el nuevo, con una "," en medio
        } else {
          if (approval.approvedBy) preDeployApprovalsApprovedByUniqueName = preDeployApprovalsApprovedByUniqueName.concat(', ', approval.approvedBy.uniqueName)
        }
        // Si es el primer item, entonces asignar el valor
        if (preDeployApprovalsModifiedDate.length === 0) {
          preDeployApprovalsModifiedDate = approval.modifiedOn;
          // Si no es el primer item, concatenar el string anterior con el nuevo, con una "," en medio
        } else {
          preDeployApprovalsModifiedDate = preDeployApprovalsModifiedDate.concat(', ', approval.modifiedOn)
        }
      }
    })
  }
  /**
   * obtener solo los valores necesarios del deployment para poblar el excel. Esto debe hacer match con el numero de columnas definidas en la constante "columns"
   */
  items.push(
    item.id,
    item.release.name,
    item.release.artifacts[0].alias,
    item.release.artifacts[0].definitionReference.definition.name,
    item.release.artifacts[0].definitionReference.repository.name,
    item.requestedFor.displayName,
    item.requestedFor.uniqueName,
    item.releaseDefinition.name,
    item.releaseEnvironment.name,
    preDeployApprovalsApproverDisplayName,
    preDeployApprovalsApprovedByDisplayName,
    preDeployApprovalsApprovedByUniqueName,
    preDeployApprovalsModifiedDate
  )

  // retornar el arreglo armado
  return items;
}

const wait = ms => new Promise(
  (resolve, reject) => setTimeout(resolve, ms)
);

const getProjectsWithDeployments = async (projects) => {
  let validProjects = [];

  try {
    for (let project of projects) {
      const deployment = await getDeployments(project.projectId, personalAccessToken)
      if (deployment.length > 0) validProjects.push({
        projectId: project.projectId,
        projectName: project.projectName,
        deployments: deployment
      })
      await wait(2000);
    }
    return validProjects;
  } catch (error) {
    console.log("an unexpected error happened", error)
    throw error;
  }

}

/**
 * Esta funcion ejecuta la escritura del archivo Excel en el mismo directorio
 */
const writeFile = async () => {

  try {

    const projects = await getProjects(personalAccessToken)
    const validProjects = await getProjectsWithDeployments(projects)

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Carlos Ramirez Vera';
    workbook.lastModifiedBy = 'Carlos Ramirez Vera';
    workbook.created = new Date(2022, 11, 11);
    workbook.modified = new Date();
    workbook.lastPrinted = new Date(2012, 11, 11);

    let i = 0;

    for (let project of validProjects) {
      const sheet = workbook.addWorksheet(project.projectName);
      sheet.properties.defaultColWidth = 32
      sheet.getColumn('A').width = 8
      sheet.getColumn('B').width = 8
      let rows = []
      for (let item of project.deployments) {
        rows.push(getRowFromItem(item))
      }
      i++
      sheet.addTable({
        name: "table" + i,
        ref: 'B2',
        headerRow: true,
        totalsRow: false,
        style: {
          theme: 'TableStyleMedium6',
          showRowStripes: true,
        },
        columns,
        rows,
      });
    }

    await workbook.xlsx.writeFile("report.xlsx");

  } catch (error) {
    console.log("ha ocurrido un error")
    console.log(error)
  }

}

writeFile();
