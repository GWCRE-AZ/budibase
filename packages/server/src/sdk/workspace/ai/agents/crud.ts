import {
  context,
  docIds,
  encryption,
  events,
  HTTPError,
} from "@budibase/backend-core"
import { DocumentType } from "@budibase/types"
import type { Agent, Optional } from "@budibase/types"
import { helpers } from "@budibase/shared-core"
import * as knowledgeBaseSdk from "../knowledgeBase"
import { v4 } from "uuid"

const SECRET_MASK = "********"
const SECRET_ENCODING_PREFIX = "bbai_enc::"
const NAME_REQUIRED_ERROR = "Agent name is required."
const DEFAULT_OPERATION_NAME = "Default operation"

type AgentLegacyFields = Partial<{
  promptInstructions: string
  enabledTools: string[]
  knowledgeBases: string[]
  knowledgeSources: NonNullable<Agent["operations"]>[number]["knowledgeSources"]
}>

type DeprecatedAgent = Agent & AgentLegacyFields

const guardName = async (name: string, id?: string) => {
  if (!name.trim()) {
    throw new HTTPError(NAME_REQUIRED_ERROR, 400)
  }

  const agents = await fetch()
  const normalizedName = helpers.normalizeForComparison(name)
  const duplicate = agents.find(
    agent =>
      helpers.normalizeForComparison(agent.name) === normalizedName &&
      agent._id !== id
  )

  if (duplicate) {
    throw new HTTPError(`Agent with name '${name}' already exists.`, 400)
  }
}

const encodeSecret = (value?: string): string | undefined => {
  if (!value || value.startsWith(SECRET_ENCODING_PREFIX)) {
    return value
  }
  return `${SECRET_ENCODING_PREFIX}${encryption.encrypt(value)}`
}

const decodeSecret = (value?: string): string | undefined => {
  if (!value || !value.startsWith(SECRET_ENCODING_PREFIX)) {
    return value
  }
  return encryption.decrypt(value.slice(SECRET_ENCODING_PREFIX.length))
}

const encodeDiscordIntegrationSecrets = (
  discordIntegration?: Agent["discordIntegration"]
) => {
  if (!discordIntegration) {
    return discordIntegration
  }

  return {
    ...discordIntegration,
    publicKey: encodeSecret(discordIntegration.publicKey),
    botToken: encodeSecret(discordIntegration.botToken),
  }
}

const decodeDiscordIntegrationSecrets = (
  discordIntegration?: Agent["discordIntegration"]
) => {
  if (!discordIntegration) {
    return discordIntegration
  }

  return {
    ...discordIntegration,
    publicKey: decodeSecret(discordIntegration.publicKey),
    botToken: decodeSecret(discordIntegration.botToken),
  }
}

const encodeSlackIntegrationSecrets = (
  slackIntegration?: Agent["slackIntegration"]
) => {
  if (!slackIntegration) {
    return slackIntegration
  }

  return {
    ...slackIntegration,
    botToken: encodeSecret(slackIntegration.botToken),
    signingSecret: encodeSecret(slackIntegration.signingSecret),
  }
}

const decodeSlackIntegrationSecrets = (
  slackIntegration?: Agent["slackIntegration"]
) => {
  if (!slackIntegration) {
    return slackIntegration
  }

  return {
    ...slackIntegration,
    botToken: decodeSecret(slackIntegration.botToken),
    signingSecret: decodeSecret(slackIntegration.signingSecret),
  }
}

const hasLegacyOperationData = (agent: AgentLegacyFields) =>
  !!agent.promptInstructions ||
  !!agent.enabledTools?.length ||
  !!agent.knowledgeBases?.length ||
  !!agent.knowledgeSources?.length

const normalizeOperation = (
  operation: NonNullable<Agent["operations"]>[number],
  fallbackId?: string
): NonNullable<Agent["operations"]>[number] => ({
  ...operation,
  id: operation.id || fallbackId || `operation_${v4()}`,
  name: operation.name?.trim() || DEFAULT_OPERATION_NAME,
  enabledTools: operation.enabledTools || [],
  knowledgeBases: operation.knowledgeBases || [],
})

const operationFromLegacyFields = (
  agent: AgentLegacyFields & Pick<Agent, "_id">
): NonNullable<Agent["operations"]>[number] => {
  const fallbackId = agent._id
    ? `operation_${helpers
        .normalizeForComparison(agent._id)
        .replace(/[^a-z0-9]/g, "_")}`
    : `operation_${v4()}`

  return normalizeOperation(
    {
      id: fallbackId,
      name: DEFAULT_OPERATION_NAME,
      promptInstructions: agent.promptInstructions,
      enabledTools: agent.enabledTools || [],
      knowledgeBases: agent.knowledgeBases || [],
      knowledgeSources: agent.knowledgeSources,
    },
    fallbackId
  )
}

const normalizeOperations = (
  agent: Pick<Agent, "_id" | "operations"> & AgentLegacyFields
) => {
  if (agent.operations?.length) {
    return agent.operations.map(operation => normalizeOperation(operation))
  }
  if (hasLegacyOperationData(agent)) {
    return [operationFromLegacyFields(agent)]
  }
  return []
}

const toStorageAgentShape = (agent: DeprecatedAgent): DeprecatedAgent => ({
  ...agent,
  operations: normalizeOperations(agent),
  promptInstructions: undefined,
  enabledTools: undefined,
  knowledgeBases: undefined,
  knowledgeSources: undefined,
})

const withAgentDefaults = (agent: DeprecatedAgent): Agent => ({
  ...agent,
  live: agent.live ?? false,
  operations: normalizeOperations(agent),
  discordIntegration: decodeDiscordIntegrationSecrets(agent.discordIntegration),
  slackIntegration: decodeSlackIntegrationSecrets(agent.slackIntegration),
})

const mergeDiscordIntegration = ({
  existing,
  incoming,
}: {
  existing?: Agent["discordIntegration"]
  incoming?: Agent["discordIntegration"]
}) => {
  if (incoming === undefined) {
    return existing
  }
  if (!incoming) {
    return incoming
  }

  const merged = {
    ...(existing || {}),
    ...incoming,
  }

  if (incoming.publicKey === SECRET_MASK && existing?.publicKey) {
    merged.publicKey = existing.publicKey
  }

  if (incoming.botToken === SECRET_MASK && existing?.botToken) {
    merged.botToken = existing.botToken
  }

  return merged
}

const mergeMSTeamsIntegration = ({
  existing,
  incoming,
}: {
  existing?: Agent["MSTeamsIntegration"]
  incoming?: Agent["MSTeamsIntegration"]
}) => {
  if (incoming === undefined) {
    return existing
  }
  if (!incoming) {
    return incoming
  }

  const merged = {
    ...(existing || {}),
    ...incoming,
  }

  if (incoming.appPassword === SECRET_MASK && existing?.appPassword) {
    merged.appPassword = existing.appPassword
  }

  return merged
}

const mergeSlackIntegration = ({
  existing,
  incoming,
}: {
  existing?: Agent["slackIntegration"]
  incoming?: Agent["slackIntegration"]
}) => {
  if (incoming === undefined) {
    return existing
  }
  if (!incoming) {
    return incoming
  }

  const merged = {
    ...(existing || {}),
    ...incoming,
  }

  if (incoming.botToken === SECRET_MASK && existing?.botToken) {
    merged.botToken = existing.botToken
  }

  if (incoming.signingSecret === SECRET_MASK && existing?.signingSecret) {
    merged.signingSecret = existing.signingSecret
  }

  return merged
}

export async function fetch(): Promise<Agent[]> {
  const db = context.getWorkspaceDB()
  const result = await db.allDocs<DeprecatedAgent>(
    docIds.getDocParams(DocumentType.AGENT, undefined, {
      include_docs: true,
    })
  )

  return result.rows
    .map(row => row.doc)
    .filter((doc): doc is DeprecatedAgent => !!doc)
    .map(withAgentDefaults)
}

export async function getOrThrow(agentId: string | undefined): Promise<Agent> {
  if (!agentId) {
    throw new HTTPError("agentId is required", 400)
  }

  const db = context.getWorkspaceDB()

  const agent = await db.tryGet<DeprecatedAgent>(agentId)
  if (!agent) {
    throw new HTTPError("Agent not found", 404)
  }

  return withAgentDefaults(agent)
}

export async function create(
  request: Optional<
    Omit<Agent, "_id" | "_rev" | "createdAt" | "updatedAt" | "publishedAt">,
    "aiconfig"
  >
): Promise<Agent> {
  const db = context.getWorkspaceDB()
  const now = new Date().toISOString()

  await guardName(request.name)

  const agent: Agent = {
    _id: docIds.generateAgentID(),
    name: request.name,
    description: request.description,
    aiconfig: request.aiconfig || "", // this might be set later, it will be validated on publish/usage
    live: request.live ?? false,
    publishedAt: request.live ? now : undefined,
    icon: request.icon,
    iconColor: request.iconColor,
    goal: request.goal,
    createdAt: now,
    createdBy: request.createdBy,
    discordIntegration: request.discordIntegration,
    MSTeamsIntegration: request.MSTeamsIntegration,
    slackIntegration: request.slackIntegration,
    operations: request.operations || [],
  }

  const storageAgent = toStorageAgentShape(agent)
  const { rev } = await db.put({
    ...storageAgent,
    discordIntegration: encodeDiscordIntegrationSecrets(
      storageAgent.discordIntegration
    ),
    slackIntegration: encodeSlackIntegrationSecrets(
      storageAgent.slackIntegration
    ),
  })
  storageAgent._rev = rev
  const result = withAgentDefaults(storageAgent)
  events.ai.agentCreated(result)
  return result
}

export async function duplicate(
  source: Agent,
  createdBy: string
): Promise<Agent> {
  const allAgents = await fetch()
  const name = helpers.duplicateName(
    source.name,
    allAgents.map(agent => agent.name)
  )

  return await create({
    name,
    description: source.description,
    aiconfig: source.aiconfig,
    goal: source.goal,
    icon: source.icon,
    iconColor: source.iconColor,
    live: source.live,
    _deleted: false,
    createdBy,
    operations: (source.operations || []).map(operation => ({
      ...operation,
      id: `operation_${v4()}`,
    })),
  })
}

export async function update(agent: Agent): Promise<Agent> {
  const { _id, _rev } = agent
  if (!_id || !_rev) {
    throw new HTTPError("_id and _rev are required", 400)
  }

  const db = context.getWorkspaceDB()
  const existingRaw = await db.tryGet<DeprecatedAgent>(_id)
  const existing = existingRaw ? withAgentDefaults(existingRaw) : undefined
  if (!existing) {
    throw new HTTPError("Agent not found", 404)
  }

  const incomingName = agent.name ?? existing.name
  const normalizedName = helpers.normalizeForComparison(incomingName)
  const normalizedExistingName = helpers.normalizeForComparison(existing.name)

  if (normalizedName !== normalizedExistingName) {
    await guardName(incomingName, _id)
  }

  const now = new Date().toISOString()
  const updated: DeprecatedAgent = {
    ...existing,
    ...agent,
    updatedAt: now,
    operations: agent.operations ?? existing?.operations,
    discordIntegration: mergeDiscordIntegration({
      existing: existing?.discordIntegration,
      incoming: agent.discordIntegration,
    }),
    MSTeamsIntegration: mergeMSTeamsIntegration({
      existing: existing?.MSTeamsIntegration,
      incoming: agent.MSTeamsIntegration,
    }),
    slackIntegration: mergeSlackIntegration({
      existing: existing?.slackIntegration,
      incoming: agent.slackIntegration,
    }),
  }

  const hasBeenPublished =
    !!existing?.publishedAt || existing?.live === true || updated.live === true
  updated.publishedAt = hasBeenPublished
    ? existing?.publishedAt || now
    : undefined

  const storageAgent = toStorageAgentShape(updated)
  const { rev } = await db.put({
    ...storageAgent,
    discordIntegration: encodeDiscordIntegrationSecrets(
      storageAgent.discordIntegration
    ),
    slackIntegration: encodeSlackIntegrationSecrets(
      storageAgent.slackIntegration
    ),
  })
  storageAgent._rev = rev
  const result = withAgentDefaults(storageAgent)
  events.ai.agentUpdated(result)
  return result
}

export async function remove(agentId: string) {
  const db = context.getWorkspaceDB()
  const agent = await getOrThrow(agentId)

  const primaryOperation = agent.operations?.[0]
  if (primaryOperation?.knowledgeBases) {
    for (const knowledgeBaseId of primaryOperation.knowledgeBases) {
      const knowledgeBase = await knowledgeBaseSdk.find(knowledgeBaseId)
      if (!knowledgeBase) {
        continue
      }

      const files =
        await knowledgeBaseSdk.listKnowledgeBaseFiles(knowledgeBaseId)
      for (const file of files) {
        try {
          await knowledgeBaseSdk.removeKnowledgeBaseFile(knowledgeBase, file)
        } catch (error) {
          console.log(
            "Failed to remove knowledge base file for agent deletion",
            {
              agentId,
              knowledgeBaseId,
              fileId: file._id,
              error,
            }
          )
        }
      }

      try {
        await knowledgeBaseSdk.remove(knowledgeBaseId)
      } catch (error) {
        console.log("Failed to remove knowledge base for agent deletion", {
          agentId,
          knowledgeBaseId,
          error,
        })
      }
    }
  }

  await db.remove(agent)
  events.ai.agentDeleted(agent)
}
