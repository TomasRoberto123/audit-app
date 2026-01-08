// SharePoint Integration Module
// Configuração
const SHAREPOINT_CONFIG = {
  siteUrl: "https://mrtauditores.sharepoint.com/sites/IT_Manager",
  lists: {
    users: "AuditUsers", // Lista de utilizadores e roles
    audits: "AuditHistory", // Histórico de auditorias
  },
  libraries: {
    documents: "AuditDocuments", // Document Library para PDFs e CSVs
  },
};

// Estado global
let currentUser = null;
let userRole = null;
let accessToken = null;

// ========== AUTENTICAÇÃO MICROSOFT ==========

/**
 * Inicializa a autenticação Microsoft usando MSAL ou SharePoint context
 */
async function initializeAuth() {
  try {
    // Primeiro, tentar usar o contexto SharePoint existente (quando aberto a partir do Power Apps)
    if (await initializeSharePointAuth()) {
      return true;
    }

    // Se não funcionar, tentar MSAL (se disponível)
    if (typeof msal !== "undefined") {
      try {
        // Nota: Para produção, deves registar uma App no Azure AD e usar o Client ID próprio
        // Por agora, vamos usar autenticação SharePoint direta
        console.log("MSAL disponível, mas a usar autenticação SharePoint direta");
      } catch (error) {
        console.warn("Erro ao inicializar MSAL:", error);
      }
    }

    return false;
  } catch (error) {
    console.error("Erro ao inicializar autenticação:", error);
    return false;
  }
}

/**
 * Autenticação alternativa usando SharePoint REST API (quando já autenticado no Power Apps)
 * Quando abres a partir do Power Apps/SharePoint, já estás autenticado via cookies
 */
async function initializeSharePointAuth() {
  try {
    // Tentar obter informações do utilizador diretamente (usa cookies de sessão)
    const user = await getUserInfo();
    if (user) {
      // Se conseguimos obter o utilizador, significa que já estamos autenticados
      return true;
    }
    return false;
  } catch (error) {
    console.error("Erro na autenticação SharePoint:", error);
    return false;
  }
}

/**
 * Obtém FormDigestValue para requisições POST (quando já autenticado)
 */
async function getFormDigestValue() {
  try {
    const response = await fetch(
      `${SHAREPOINT_CONFIG.siteUrl}/_api/contextinfo`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
        },
        credentials: "include", // Importante: incluir cookies de autenticação
      }
    );

    if (response.ok) {
      const data = await response.json();
      return data.d.GetContextWebInformation.FormDigestValue;
    }
  } catch (error) {
    console.error("Erro ao obter FormDigestValue:", error);
  }
  return null;
}

/**
 * Obtém informações do utilizador atual
 * Quando aberto a partir do Power Apps/SharePoint, usa os cookies de sessão automaticamente
 */
async function getUserInfo() {
  try {
    // Quando já estás autenticado no Power Apps/SharePoint, os cookies são enviados automaticamente
    // Não precisamos de token Bearer - a sessão já está autenticada
    const response = await fetch(
      `${SHAREPOINT_CONFIG.siteUrl}/_api/web/currentuser`,
      {
        headers: {
          Accept: "application/json;odata=verbose",
        },
        credentials: "include", // CRUCIAL: incluir cookies de autenticação
      }
    );

    if (response.ok) {
      const data = await response.json();
      currentUser = {
        email: data.d.Email,
        name: data.d.Title,
        loginName: data.d.LoginName,
      };
      await loadUserRole();
      return currentUser;
    } else if (response.status === 401) {
      // Não autenticado - pode precisar de login
      console.warn("Não autenticado. Pode ser necessário fazer login.");
      return null;
    }
  } catch (error) {
    console.error("Erro ao obter informações do utilizador:", error);
  }
  return null;
}

/**
 * Carrega o role do utilizador da lista SharePoint
 */
async function loadUserRole() {
  try {
    if (!currentUser) return;

    const userEmail = currentUser.email || currentUser.loginName;
    const users = await getListItems(SHAREPOINT_CONFIG.lists.users, {
      filter: `Email eq '${userEmail}'`,
    });

    if (users.length > 0) {
      userRole = users[0].Role || "Empregado";
    } else {
      // Se não encontrar, criar registo com role padrão "Empregado"
      userRole = "Empregado";
      await createListItem(SHAREPOINT_CONFIG.lists.users, {
        Email: userEmail,
        Nome: currentUser.name || userEmail,
        Role: "Empregado",
      });
    }
  } catch (error) {
    console.error("Erro ao carregar role:", error);
    userRole = "Empregado"; // Default
  }
}

/**
 * Verifica se o utilizador é chefe
 */
function isChefe() {
  return userRole === "Chefe";
}

/**
 * Obtém o email do utilizador atual
 */
function getCurrentUserEmail() {
  return currentUser?.email || currentUser?.loginName || "";
}

// ========== SHAREPOINT REST API ==========

/**
 * Faz uma chamada à REST API do SharePoint
 * Usa cookies de sessão quando já estás autenticado
 */
async function sharePointRequest(endpoint, options = {}) {
  try {
    const url = endpoint.startsWith("http")
      ? endpoint
      : `${SHAREPOINT_CONFIG.siteUrl}/_api${endpoint}`;

    const defaultHeaders = {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
    };

    // Para requisições POST, precisamos do FormDigestValue
    if (options.method === "POST" && !options.headers?.["X-RequestDigest"]) {
      const digest = await getFormDigestValue();
      if (digest) {
        defaultHeaders["X-RequestDigest"] = digest;
      }
    }

    const response = await fetch(url, {
      ...options,
      credentials: "include", // CRUCIAL: incluir cookies de autenticação
      headers: {
        ...defaultHeaders,
        ...options.headers,
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`SharePoint API Error: ${response.status} - ${errorText}`);
    }

    return await response.json();
  } catch (error) {
    console.error("Erro na chamada SharePoint:", error);
    throw error;
  }
}

/**
 * Obtém itens de uma lista SharePoint
 */
async function getListItems(listName, options = {}) {
  try {
    let url = `/web/lists/getbytitle('${listName}')/items`;
    const params = [];

    if (options.filter) {
      params.push(`$filter=${encodeURIComponent(options.filter)}`);
    }
    if (options.select) {
      params.push(`$select=${encodeURIComponent(options.select)}`);
    }
    if (options.orderBy) {
      params.push(`$orderby=${encodeURIComponent(options.orderBy)}`);
    }
    if (options.top) {
      params.push(`$top=${options.top}`);
    }

    if (params.length > 0) {
      url += "?" + params.join("&");
    }

    const response = await sharePointRequest(url);
    return response.d.results.map((item) => ({
      id: item.Id,
      ...item,
    }));
  } catch (error) {
    console.error(`Erro ao obter itens da lista ${listName}:`, error);
    return [];
  }
}

/**
 * Cria um item numa lista SharePoint
 */
async function createListItem(listName, fields) {
  try {
    const url = `/web/lists/getbytitle('${listName}')/items`;
    const payload = {
      __metadata: {
        type: `SP.Data.${listName.replace(/\s/g, "_x0020_")}ListItem`,
      },
      ...fields,
    };

    const response = await sharePointRequest(url, {
      method: "POST",
      body: JSON.stringify(payload),
    });

    return response.d;
  } catch (error) {
    console.error(`Erro ao criar item na lista ${listName}:`, error);
    throw error;
  }
}

/**
 * Atualiza um item numa lista SharePoint
 */
async function updateListItem(listName, itemId, fields) {
  try {
    const url = `/web/lists/getbytitle('${listName}')/items(${itemId})`;
    const payload = {
      __metadata: {
        type: `SP.Data.${listName.replace(/\s/g, "_x0020_")}ListItem`,
      },
      ...fields,
    };

    const response = await sharePointRequest(url, {
      method: "POST",
      headers: {
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      },
      body: JSON.stringify(payload),
    });

    return response;
  } catch (error) {
    console.error(`Erro ao atualizar item na lista ${listName}:`, error);
    throw error;
  }
}

/**
 * Elimina um item de uma lista SharePoint
 */
async function deleteListItem(listName, itemId) {
  try {
    const url = `/web/lists/getbytitle('${listName}')/items(${itemId})`;

    await sharePointRequest(url, {
      method: "POST",
      headers: {
        "X-HTTP-Method": "DELETE",
        "IF-MATCH": "*",
      },
    });

    return true;
  } catch (error) {
    console.error(`Erro ao eliminar item da lista ${listName}:`, error);
    throw error;
  }
}

/**
 * Faz upload de um ficheiro para uma Document Library
 */
async function uploadFileToLibrary(libraryName, fileName, fileContent, contentType = "application/pdf") {
  try {
    const digestValue = await getFormDigestValue();
    if (!digestValue) {
      throw new Error("Não foi possível obter FormDigestValue");
    }

    // Converter fileContent para ArrayBuffer se necessário
    let arrayBuffer;
    if (fileContent instanceof Blob) {
      arrayBuffer = await fileContent.arrayBuffer();
    } else if (fileContent instanceof ArrayBuffer) {
      arrayBuffer = fileContent;
    } else {
      // Assumir que é base64 ou string
      if (typeof fileContent === "string" && fileContent.startsWith("data:")) {
        const base64 = fileContent.split(",")[1];
        const binaryString = atob(base64);
        arrayBuffer = new ArrayBuffer(binaryString.length);
        const uint8Array = new Uint8Array(arrayBuffer);
        for (let i = 0; i < binaryString.length; i++) {
          uint8Array[i] = binaryString.charCodeAt(i);
        }
      } else {
        throw new Error("Formato de ficheiro não suportado");
      }
    }

    const url = `${SHAREPOINT_CONFIG.siteUrl}/_api/web/getfolderbyserverrelativeurl('${libraryName}')/files/add(url='${fileName}', overwrite=true)`;

    const response = await fetch(url, {
      method: "POST",
      headers: {
        "X-RequestDigest": digestValue,
        Accept: "application/json;odata=verbose",
      },
      credentials: "include", // CRUCIAL: incluir cookies de autenticação
      body: arrayBuffer,
    });

    if (!response.ok) {
      throw new Error(`Erro no upload: ${response.status}`);
    }

    const result = await response.json();
    return result.d;
  } catch (error) {
    console.error(`Erro ao fazer upload para ${libraryName}:`, error);
    throw error;
  }
}

// ========== FUNÇÕES ESPECÍFICAS PARA AUDITORIA ==========

/**
 * Guarda uma auditoria no SharePoint
 */
async function saveAuditToSharePoint(sections, context, metadata) {
  try {
    const userEmail = getCurrentUserEmail();
    const auditData = {
      sections,
      context,
      metadata,
    };

    // Criar registo na lista
    const listItem = await createListItem(SHAREPOINT_CONFIG.lists.audits, {
      Title: metadata.adjudicantes[0] || "Não especificado",
      Entidade: metadata.adjudicantes[0] || "Não especificado",
      Auditor: metadata.auditorName,
      DataEmissao: metadata.reportDate,
      CriadoPor: userEmail,
      TotalContratos: metadata.totalContracts,
      TotalAchados: Object.values(sections).reduce(
        (sum, section) => sum + (section.findings?.length || 0),
        0
      ),
      AnoN: context.yearN,
      AnoN1: context.yearNMinus1 || "",
      AnoN2: context.yearNMinus2 || "",
      DadosJSON: JSON.stringify(auditData),
    });

    // Gerar PDF e fazer upload
    try {
      // Nota: generatePdf precisa ser chamado de app.js
      // Por agora, guardamos apenas os dados
      // O PDF pode ser gerado quando necessário
    } catch (pdfError) {
      console.error("Erro ao gerar PDF:", pdfError);
    }

    return listItem;
  } catch (error) {
    console.error("Erro ao guardar auditoria no SharePoint:", error);
    throw error;
  }
}

/**
 * Carrega histórico de auditorias do SharePoint
 */
async function loadAuditHistoryFromSharePoint() {
  try {
    const userEmail = getCurrentUserEmail();
    let filter = "";

    // Se não for chefe, filtrar apenas os seus registos
    if (!isChefe()) {
      filter = `CriadoPor eq '${userEmail}'`;
    }

    const items = await getListItems(SHAREPOINT_CONFIG.lists.audits, {
      filter,
      orderBy: "DataEmissao desc",
      top: 100,
    });

    return items.map((item) => ({
      id: item.id.toString(),
      entity: item.Entidade || item.Title,
      auditor: item.Auditor,
      date: item.DataEmissao,
      timestamp: new Date(item.Created).getTime(),
      totalContracts: item.TotalContratos,
      totalFindings: item.TotalAchados,
      years: {
        n: item.AnoN,
        n1: item.AnoN1,
        n2: item.AnoN2,
      },
      data: JSON.parse(item.DadosJSON || "{}"),
      createdBy: item.CriadoPor,
    }));
  } catch (error) {
    console.error("Erro ao carregar histórico do SharePoint:", error);
    return [];
  }
}

/**
 * Elimina uma auditoria do SharePoint
 */
async function deleteAuditFromSharePoint(auditId) {
  try {
    const userEmail = getCurrentUserEmail();
    
    // Verificar se o utilizador pode eliminar (só se for chefe ou se for o criador)
    const items = await getListItems(SHAREPOINT_CONFIG.lists.audits, {
      filter: `Id eq ${auditId}`,
    });

    if (items.length === 0) {
      throw new Error("Auditoria não encontrada");
    }

    const audit = items[0];
    if (!isChefe() && audit.CriadoPor !== userEmail) {
      throw new Error("Não tem permissão para eliminar esta auditoria");
    }

    await deleteListItem(SHAREPOINT_CONFIG.lists.audits, auditId);
    return true;
  } catch (error) {
    console.error("Erro ao eliminar auditoria:", error);
    throw error;
  }
}

/**
 * Obtém uma auditoria específica do SharePoint
 */
async function getAuditFromSharePoint(auditId) {
  try {
    const items = await getListItems(SHAREPOINT_CONFIG.lists.audits, {
      filter: `Id eq ${auditId}`,
    });

    if (items.length === 0) {
      return null;
    }

    const item = items[0];
    const userEmail = getCurrentUserEmail();

    // Verificar permissões
    if (!isChefe() && item.CriadoPor !== userEmail) {
      throw new Error("Não tem permissão para ver esta auditoria");
    }

    return {
      id: item.id.toString(),
      entity: item.Entidade || item.Title,
      auditor: item.Auditor,
      date: item.DataEmissao,
      timestamp: new Date(item.Created).getTime(),
      totalContracts: item.TotalContratos,
      totalFindings: item.TotalAchados,
      years: {
        n: item.AnoN,
        n1: item.AnoN1,
        n2: item.AnoN2,
      },
      data: JSON.parse(item.DadosJSON || "{}"),
      createdBy: item.CriadoPor,
    };
  } catch (error) {
    console.error("Erro ao obter auditoria:", error);
    throw error;
  }
}

// Exportar funções principais
window.SharePointIntegration = {
  initialize: initializeAuth,
  isChefe,
  getCurrentUserEmail,
  saveAudit: saveAuditToSharePoint,
  loadHistory: loadAuditHistoryFromSharePoint,
  deleteAudit: deleteAuditFromSharePoint,
  getAudit: getAuditFromSharePoint,
  uploadFile: uploadFileToLibrary,
  config: SHAREPOINT_CONFIG,
};

