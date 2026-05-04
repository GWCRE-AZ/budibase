const mockCacheStore = jest.fn()
const mockGetPlatformUrl = jest.fn()
const mockGetOrThrowWorkspaceId = jest.fn()
const mockNewId = jest.fn()
const mockSetCookie = jest.fn()

jest.mock("@budibase/backend-core", () => ({
  cache: {
    store: (...args: any[]) => mockCacheStore(...args),
  },
  configs: {
    getPlatformUrl: (...args: any[]) => mockGetPlatformUrl(...args),
  },
  constants: {
    Cookie: {
      DatasourceAuth: "datasourceAuth",
    },
  },
  context: {
    getOrThrowWorkspaceId: (...args: any[]) => mockGetOrThrowWorkspaceId(...args),
  },
  env: {
    MICROSOFT_CLIENT_ID: "ms-client-id",
    MICROSOFT_CLIENT_SECRET: "ms-client-secret",
    MICROSOFT_TENANT_ID: "common",
    RAG_SHAREPOINT_DEFAULT_SCOPE: "openid profile offline_access",
  },
  utils: {
    newid: (...args: any[]) => mockNewId(...args),
    setCookie: (...args: any[]) => mockSetCookie(...args),
  },
}))

import { startSharePointAuth } from "../sharepointAuth"

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
        returnPath: "/builder/workspace/app_dev_trusted_workspace/agent/abc/knowledge",
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
      "datasourceAuth"
    )
    expect(redirect).toHaveBeenCalledTimes(1)
    expect(redirect.mock.calls[0][0]).toContain(
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    )
    expect(redirect.mock.calls[0][0]).toContain("state=state_123")
  })
})
