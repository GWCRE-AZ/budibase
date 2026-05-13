import {
  AgentFileUploadResponse,
  ConnectAgentSharePointSiteRequest,
  ConnectAgentSharePointSiteResponse,
  CreateAgentRequest,
  CreateAgentResponse,
  DisconnectAgentSharePointSiteResponse,
  DuplicateAgentResponse,
  FetchAgentKnowledgeResponse,
  FetchAgentKnowledgeSourceEntriesResponse,
  FetchAgentKnowledgeSourceOptionsResponse,
  FetchAgentsResponse,
  ProvisionAgentSlackChannelRequest,
  ProvisionAgentSlackChannelResponse,
  ProvisionAgentMSTeamsChannelRequest,
  ProvisionAgentMSTeamsChannelResponse,
  SyncAgentDiscordCommandsRequest,
  SyncAgentDiscordCommandsResponse,
  SyncAgentKnowledgeSourcesRequest,
  SyncAgentKnowledgeSourcesResponse,
  ToggleAgentDeploymentRequest,
  ToggleAgentDeploymentResponse,
  ToolMetadata,
  UpdateAgentSharePointSiteRequest,
  UpdateAgentSharePointSiteResponse,
  UpdateAgentRequest,
  UpdateAgentResponse,
} from "@budibase/types"

import { BaseAPIClient } from "./types"

export interface AgentEndpoints {
  fetchTools: (aiconfigId?: string) => Promise<ToolMetadata[]>
  fetchAgents: () => Promise<FetchAgentsResponse>
  createAgent: (agent: CreateAgentRequest) => Promise<CreateAgentResponse>
  updateAgent: (agent: UpdateAgentRequest) => Promise<UpdateAgentResponse>
  duplicateAgent: (agentId: string) => Promise<DuplicateAgentResponse>
  deleteAgent: (agentId: string) => Promise<{ deleted: true }>
  syncAgentDiscordCommands: (
    agentId: string,
    body?: SyncAgentDiscordCommandsRequest
  ) => Promise<SyncAgentDiscordCommandsResponse>
  provisionAgentMSTeamsChannel: (
    agentId: string,
    body?: ProvisionAgentMSTeamsChannelRequest
  ) => Promise<ProvisionAgentMSTeamsChannelResponse>
  provisionAgentSlackChannel: (
    agentId: string,
    body?: ProvisionAgentSlackChannelRequest
  ) => Promise<ProvisionAgentSlackChannelResponse>
  toggleAgentDiscordDeployment: (
    agentId: string,
    enabled: boolean
  ) => Promise<ToggleAgentDeploymentResponse>
  toggleAgentMSTeamsDeployment: (
    agentId: string,
    enabled: boolean
  ) => Promise<ToggleAgentDeploymentResponse>
  toggleAgentSlackDeployment: (
    agentId: string,
    enabled: boolean
  ) => Promise<ToggleAgentDeploymentResponse>
  fetchAgentKnowledge: (
    agentId: string,
    operationId?: string
  ) => Promise<FetchAgentKnowledgeResponse>
  uploadAgentFile: (
    agentId: string,
    file: File,
    operationId?: string
  ) => Promise<AgentFileUploadResponse>
  deleteAgentFile: (
    agentId: string,
    fileId: string,
    operationId?: string
  ) => Promise<{ deleted: true }>
  fetchAgentKnowledgeSourceOptions: (
    datasourceId: string,
    authConfigId: string
  ) => Promise<FetchAgentKnowledgeSourceOptionsResponse>
  fetchAgentKnowledgeSourceAllEntries: (
    agentId: string,
    siteId: string,
    operationId?: string
  ) => Promise<FetchAgentKnowledgeSourceEntriesResponse>
  connectAgentSharePointSite: (
    agentId: string,
    body: ConnectAgentSharePointSiteRequest,
    operationId?: string
  ) => Promise<ConnectAgentSharePointSiteResponse>
  updateAgentSharePointSite: (
    agentId: string,
    siteId: string,
    body: UpdateAgentSharePointSiteRequest,
    operationId?: string
  ) => Promise<UpdateAgentSharePointSiteResponse>
  disconnectAgentSharePointSite: (
    agentId: string,
    siteId: string,
    operationId?: string
  ) => Promise<DisconnectAgentSharePointSiteResponse>
  syncAgentKnowledgeSources: (
    agentId: string,
    sourceId: string,
    operationId?: string
  ) => Promise<SyncAgentKnowledgeSourcesResponse>
}

export const buildAgentEndpoints = (API: BaseAPIClient): AgentEndpoints => ({
  fetchTools: async (aiconfigId?: string) => {
    const query = aiconfigId
      ? `?aiconfigId=${encodeURIComponent(aiconfigId)}`
      : ""
    return await API.get({
      url: `/api/agent/tools${query}`,
    })
  },
  fetchAgents: async () => {
    return await API.get({
      url: "/api/agent",
    })
  },

  createAgent: async (agent: CreateAgentRequest) => {
    return await API.post({
      url: "/api/agent",
      body: agent,
    })
  },

  updateAgent: async (agent: UpdateAgentRequest) => {
    return await API.put({
      url: "/api/agent",
      body: agent,
    })
  },

  duplicateAgent: async (agentId: string) => {
    return await API.post({
      url: `/api/agent/${agentId}/duplicate`,
    })
  },

  deleteAgent: async (agentId: string) => {
    return await API.delete({
      url: `/api/agent/${agentId}`,
    })
  },

  syncAgentDiscordCommands: async (agentId: string, body) => {
    return await API.post<
      SyncAgentDiscordCommandsRequest | undefined,
      SyncAgentDiscordCommandsResponse
    >({
      url: `/api/agent/${agentId}/discord/sync`,
      body,
    })
  },

  provisionAgentMSTeamsChannel: async (agentId: string, body) => {
    return await API.post<
      ProvisionAgentMSTeamsChannelRequest | undefined,
      ProvisionAgentMSTeamsChannelResponse
    >({
      url: `/api/agent/${agentId}/ms-teams/provision`,
      body,
    })
  },

  provisionAgentSlackChannel: async (agentId: string, body) => {
    return await API.post<
      ProvisionAgentSlackChannelRequest | undefined,
      ProvisionAgentSlackChannelResponse
    >({
      url: `/api/agent/${agentId}/slack/provision`,
      body,
    })
  },

  toggleAgentDiscordDeployment: async (agentId: string, enabled: boolean) => {
    return await API.post<
      ToggleAgentDeploymentRequest,
      ToggleAgentDeploymentResponse
    >({
      url: `/api/agent/${agentId}/discord/toggle`,
      body: { enabled },
    })
  },

  toggleAgentMSTeamsDeployment: async (agentId: string, enabled: boolean) => {
    return await API.post<
      ToggleAgentDeploymentRequest,
      ToggleAgentDeploymentResponse
    >({
      url: `/api/agent/${agentId}/ms-teams/toggle`,
      body: { enabled },
    })
  },

  toggleAgentSlackDeployment: async (agentId: string, enabled: boolean) => {
    return await API.post<
      ToggleAgentDeploymentRequest,
      ToggleAgentDeploymentResponse
    >({
      url: `/api/agent/${agentId}/slack/toggle`,
      body: { enabled },
    })
  },

  fetchAgentKnowledge: async (agentId: string, operationId?: string) => {
    const query = operationId
      ? `?operationId=${encodeURIComponent(operationId)}`
      : ""
    return await API.get<FetchAgentKnowledgeResponse>({
      url: `/api/agent/${agentId}/knowledge${query}`,
    })
  },

  uploadAgentFile: async (
    agentId: string,
    file: File,
    operationId?: string
  ) => {
    const formData = new FormData()
    formData.append("file", file)
    const query = operationId
      ? `?operationId=${encodeURIComponent(operationId)}`
      : ""
    return await API.post<FormData, AgentFileUploadResponse>({
      url: `/api/agent/${agentId}/files${query}`,
      body: formData,
      json: false,
    })
  },

  deleteAgentFile: async (
    agentId: string,
    fileId: string,
    operationId?: string
  ) => {
    const query = operationId
      ? `?operationId=${encodeURIComponent(operationId)}`
      : ""
    return await API.delete({
      url: `/api/agent/${agentId}/files/${fileId}${query}`,
    })
  },

  fetchAgentKnowledgeSourceOptions: async (
    datasourceId: string,
    authConfigId: string
  ) => {
    return await API.get<FetchAgentKnowledgeSourceOptionsResponse>({
      url: `/api/knowledge-sources/${encodeURIComponent(datasourceId)}/${encodeURIComponent(authConfigId)}/options`,
    })
  },

  fetchAgentKnowledgeSourceAllEntries: async (
    agentId: string,
    siteId: string,
    operationId?: string
  ) => {
    const query = new URLSearchParams({ siteId })
    if (operationId) {
      query.set("operationId", operationId)
    }
    return await API.get<FetchAgentKnowledgeSourceEntriesResponse>({
      url: `/api/agent/${agentId}/knowledge-sources/sharepoint/entries/all?${query.toString()}`,
    })
  },

  connectAgentSharePointSite: async (
    agentId: string,
    body,
    operationId?: string
  ) => {
    const query = operationId
      ? `?operationId=${encodeURIComponent(operationId)}`
      : ""
    return await API.post<
      ConnectAgentSharePointSiteRequest,
      ConnectAgentSharePointSiteResponse
    >({
      url: `/api/agent/${agentId}/knowledge-sources/sharepoint/sites${query}`,
      body,
    })
  },

  updateAgentSharePointSite: async (
    agentId: string,
    siteId: string,
    body,
    operationId?: string
  ) => {
    const query = operationId
      ? `?operationId=${encodeURIComponent(operationId)}`
      : ""
    return await API.patch<
      UpdateAgentSharePointSiteRequest,
      UpdateAgentSharePointSiteResponse
    >({
      url: `/api/agent/${agentId}/knowledge-sources/sharepoint/sites/${encodeURIComponent(siteId)}${query}`,
      body,
    })
  },

  disconnectAgentSharePointSite: async (
    agentId: string,
    siteId: string,
    operationId?: string
  ) => {
    const query = operationId
      ? `?operationId=${encodeURIComponent(operationId)}`
      : ""
    return await API.delete<void, DisconnectAgentSharePointSiteResponse>({
      url: `/api/agent/${agentId}/knowledge-sources/sharepoint/sites/${encodeURIComponent(siteId)}${query}`,
    })
  },

  syncAgentKnowledgeSources: async (
    agentId: string,
    sourceId: string,
    operationId?: string
  ) => {
    const query = operationId
      ? `?operationId=${encodeURIComponent(operationId)}`
      : ""
    return await API.post<
      SyncAgentKnowledgeSourcesRequest | undefined,
      SyncAgentKnowledgeSourcesResponse
    >({
      url: `/api/agent/${agentId}/knowledge-sources/${encodeURIComponent(sourceId)}/sync${query}`,
    })
  },
})
