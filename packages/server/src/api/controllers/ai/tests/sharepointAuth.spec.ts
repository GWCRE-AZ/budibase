const mockCacheStore = jest.fn()
const mockGetPlatformUrl = jest.fn()
const mockGetOrThrowWorkspaceId = jest.fn()
const mockNewId = jest.fn()
const mockSetCookie = jest.fn()

jest.mock("@budibase/backend-core", () => {
  const actual = jest.requireActual("@budibase/backend-core")
  return {
    ...actual,
    cache: {
      ...actual.cache,
      store: (...args: any[]) => mockCacheStore(...args),
    },
    configs: {
      ...actual.configs,
      getPlatformUrl: (...args: any[]) => mockGetPlatformUrl(...args),
    },
    context: {
      ...actual.context,
      getOrThrowWorkspaceId: (...args: any[]) =>
        mockGetOrThrowWorkspaceId(...args),
    },
    env: {
      ...actual.env,
      MICROSOFT_CLIENT_ID: "ms-client-id",
      MICROSOFT_CLIENT_SECRET: "ms-client-secret",
      MICROSOFT_TENANT_ID: "common",
      RAG_SHAREPOINT_DEFAULT_SCOPE: "openid profile offline_access",
    },
    utils: {
      ...actual.utils,
      newid: (...args: any[]) => mockNewId(...args),
      setCookie: (...args: any[]) => mockSetCookie(...args),
    },
  }
})

import { startSharePointAuth } from "../sharepointAuth"
import { constants } from "@budibase/backend-core"

describe("sharepointAuth controller", () => {
  beforeEach(() => {
    jest.clearAllMocks()
    mockGetOrThrowWorkspaceId.mockReturnValue("app_dev_trusted_workspace")
    mockGetPlatformUrl.mockResolvedValue("http://localhost:10000")
    mockCacheStore.mockResolvedValue(undefined)
    mockNewId.mockReturnValue("state_123")
  })

  it("uses trusted workspace context, not user-controlled query params", async () => {
    const redirect = jest.fn()
    const ctx = {
      query: {
        returnPath:
          "/builder/workspace/app_dev_trusted_workspace/agent/abc/knowledge",
        appId: "app_dev_attacker_workspace",
      },
      redirect,
    } as any

    await startSharePointAuth(ctx)

    expect(mockGetOrThrowWorkspaceId).toHaveBeenCalledTimes(1)
    expect(mockCacheStore).toHaveBeenCalledWith(
      "datasource:microsoft:state:state_123",
      {
        workspaceId: "app_dev_trusted_workspace",
        provider: "microsoft",
      },
      600
    )
    expect(mockSetCookie).toHaveBeenCalledWith(
      ctx,
      expect.objectContaining({
        workspaceId: "app_dev_trusted_workspace",
        provider: "microsoft",
        returnPath:
          "/builder/workspace/app_dev_trusted_workspace/agent/abc/knowledge",
      }),
      constants.Cookie.DatasourceAuth
    )
    expect(redirect).toHaveBeenCalledTimes(1)
    expect(redirect.mock.calls[0][0]).toContain(
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    )
    expect(redirect.mock.calls[0][0]).toContain("state=state_123")
  })

  it("ignores a forged workspaceId in query and keeps trusted context workspace", async () => {
    const redirect = jest.fn()
    const ctx = {
      query: {
        returnPath:
          "/builder/workspace/app_dev_attacker_workspace/agent/abc/knowledge",
        workspaceId: "app_dev_attacker_workspace",
      },
      redirect,
    } as any

    await startSharePointAuth(ctx)

    expect(mockCacheStore).toHaveBeenCalledWith(
      "datasource:microsoft:state:state_123",
      {
        workspaceId: "app_dev_trusted_workspace",
        provider: "microsoft",
      },
      600
    )
    expect(mockSetCookie).toHaveBeenCalledWith(
      ctx,
      expect.objectContaining({
        workspaceId: "app_dev_trusted_workspace",
      }),
      constants.Cookie.DatasourceAuth
    )
    expect(redirect).toHaveBeenCalledTimes(1)
  })
})
