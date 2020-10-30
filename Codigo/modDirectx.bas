Attribute VB_Name = "modDirectx"
Option Explicit
Public SurfaceDB As New clsTextureManager
Public Batch As New clsSpriteBatch

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Declare Function video_create_device Lib "utils.dll" (ByVal window As Long, ByVal width As Long, ByVal height As Long, ByVal vsync As Long) As Long
Public Declare Sub video_cleanup_device Lib "utils.dll" ()

Public Declare Sub video_clear_color Lib "utils.dll" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
Public Declare Sub video_clear Lib "utils.dll" (ByVal buffer As Long)

Public Declare Sub video_begin_draw Lib "utils.dll" ()
Public Declare Sub video_end_draw Lib "utils.dll" ()

Public Declare Sub video_swap_buffers Lib "utils.dll" ()
Public Declare Sub video_swap_buffers_to Lib "utils.dll" (dst As RECT, ByVal hwnd As Long)

Public Declare Function video_create_texture_from_file Lib "utils.dll" (ByVal filename As String) As Long
Public Declare Function video_create_texture Lib "utils.dll" (ByVal width As Long, ByVal height As Long, ByVal renderTarget As Long) As Long

Public Declare Sub video_set_render_target Lib "utils.dll" (ByVal width As Long)

Public Declare Sub video_draw_primitive Lib "utils.dll" (ByVal ptype As Long, ByVal start As Long, ByVal primCount As Long)
Public Declare Sub video_draw_indexed_primitive Lib "utils.dll" (ByVal ptype As Long, ByVal numVertex As Long, ByVal start As Long, ByVal primCount As Long)

Public Declare Sub video_set_index_buffer Lib "utils.dll" (ByVal handle As Long)
Public Declare Sub video_set_vertex_buffer Lib "utils.dll" (ByVal handle As Long)
Public Declare Sub video_set_shader Lib "utils.dll" (ByVal handle As Long)
Public Declare Sub video_set_texture Lib "utils.dll" (ByVal handle As Long, ByVal location As Long)

Public Declare Function video_create_vertex_buffer Lib "utils.dll" (ByVal size As Long, data As Any, ByVal usage As Long) As Long
Public Declare Function video_create_index_buffer Lib "utils.dll" (ByVal size As Long, data As Any, ByVal format As Long, ByVal usage As Long) As Long
Public Declare Function video_create_shader_from_filename Lib "utils.dll" (ByVal filename As String) As Long

Declare Sub video_erase_vertex_buffer Lib "utils.dll" (ByVal handle As Long)
Declare Sub video_erase_index_buffer Lib "utils.dll" (ByVal handle As Long)
Declare Sub video_erase_texture Lib "utils.dll" (ByVal handle As Long)

Public Declare Sub video_set_shader_paramater_float Lib "utils.dll" (ByVal handle As Long, ByVal name As String, ByVal value As Single)
Public Declare Sub video_set_shader_paramater_float2 Lib "utils.dll" (ByVal handle As Long, ByVal name As String, value As Vec2)
Public Declare Sub video_set_shader_paramater_float3 Lib "utils.dll" (ByVal handle As Long, ByVal name As String, value As Vec3)
Public Declare Sub video_set_shader_paramater_float4 Lib "utils.dll" (ByVal handle As Long, ByVal name As String, value As Vec4)
Public Declare Sub video_set_shader_paramater_matrix Lib "utils.dll" (ByVal handle As Long, ByVal name As String, value As Matrix)
Public Declare Sub video_set_shader_technique Lib "utils.dll" (ByVal handle As Long, ByVal name As String)

Declare Sub video_update_vertex_buffer Lib "utils.dll" (ByVal handle As Long, vPtr As Any, ByVal size As Long)
Declare Sub video_update_index_buffer Lib "utils.dll" (ByVal handle As Long, vPtr As Any, ByVal size As Long)

Public Declare Function video_get_texture_info Lib "utils.dll" (ByVal handle As Long, width As Long, height As Long) As Long

Public Const D3DFMT_INDEX16 = 101
Public Const D3DFMT_INDEX32 = 102

Public Const D3DUSAGE_DYNAMIC = 512
Public Const D3DUSAGE_WRITEONLY = 8

Public Const D3DPT_POINTLIST = 1
Public Const D3DPT_LINELIST = 2
Public Const D3DPT_LINESTRIP = 3
Public Const D3DPT_TRIANGLELIST = 4
Public Const D3DPT_TRIANGLESTRIP = 5
Public Const D3DPT_TRIANGLEFAN = 6

Public ScreenWidth As Long
Public ScreenHeight As Long

Public MainScreenRect As RECT
Public FPS As Long

Public ShaderHandler As Long
Public SceneTexture As Long

Public Function Initialize(ByVal hwnd As Long, ByVal width As Long, ByVal height As Long, ByVal vsync As Long) As Boolean
'***************************************************
'
'
'***************************************************
    ScreenWidth = width
    ScreenHeight = height

    If video_create_device(hwnd, width, height, vsync) = 0 Then
        Call MsgBox("Error al iniciar dx9")
        Call CloseClient
        Exit Function
    End If
    
    SceneTexture = video_create_texture(width, height, 1)

    If Not Batch.Initialize(1024) And SceneTexture <> -1 Then
        Call MsgBox("Error al iniciar Batch")
        Call CloseClient
        Exit Function
    End If
    
    ShaderHandler = CreateShaderFromFile(App.path & "\Content\Shaders\basic.fx")
    If ShaderHandler = -1 Then
        Call MsgBox("Error al iniciar shader")
        Call CloseClient
        Exit Function
    End If

    With MainScreenRect
        .bottom = ScreenHeight
        .right = ScreenWidth
    End With

    Engine_Init_FontTextures
    Engine_Init_FontSettings


    Dim view As Matrix
    Call matrix_ortho_off_centerLH(view, 0, width, height, 0, -1#, 1#)
    
    Call SetShaderUniformMatrix(ShaderHandler, "g_view", view)
    Call SetShader(ShaderHandler)
    Call ClearColor(0, 0, 0, 255)
End Function

Public Sub Cleanup()
    Call video_cleanup_device
End Sub

Public Sub BeginScene()
'***************************************************
'
'
'***************************************************
    Call video_begin_draw

End Sub

Public Sub EndScene()
'***************************************************
'
'
'***************************************************
    Call video_end_draw
End Sub

Public Sub SwapBuffer()
'***************************************************
'
'
'***************************************************
    Call video_swap_buffers
End Sub

Public Sub SwapBufferTo(ByVal hwnd As Long, dest As RECT)
'***************************************************
'
'
'***************************************************
    Call video_swap_buffers_to(dest, hwnd)
End Sub

Public Sub ClearColor(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
'***************************************************
'
'
'***************************************************
    Call video_clear_color(r, g, b, a)
End Sub

Public Sub ClearBuffer(ByVal buffer As Long)
'***************************************************
'
'
'***************************************************
    Call video_clear(buffer)
End Sub

Public Sub DrawPrimitive(ByVal ptype As Long, ByVal start As Long, ByVal primCount As Long)
'***************************************************
'
'
'***************************************************
    Call video_draw_primitive(ptype, start, primCount)
End Sub

Public Sub DrawIndexedPrimitive(ByVal ptype As Long, ByVal numVertex As Long, ByVal start As Long, ByVal primCount As Long)
'***************************************************
'
'
'***************************************************
    Call video_draw_indexed_primitive(ptype, numVertex, start, primCount)
End Sub

Public Sub SetVertexBuffer(ByVal handle As Long)
'***************************************************
'
'
'***************************************************
    Call video_set_vertex_buffer(handle)
End Sub

Public Sub SetIndexBuffer(ByVal handle As Long)
'***************************************************
'
'
'***************************************************
    Call video_set_index_buffer(handle)
End Sub

Public Sub SetShader(ByVal handle As Long)
'***************************************************
'
'
'***************************************************
    Call video_set_shader(handle)
End Sub

Public Sub SetTexture(ByVal handle As Long, ByVal location As Long)
'***************************************************
'
'
'***************************************************
    Call video_set_texture(handle, location)
End Sub

Public Function CreateShaderFromFile(filename As String) As Long
'***************************************************
'
'
'***************************************************
    CreateShaderFromFile = video_create_shader_from_filename(filename)
End Function

Public Function CreateTextureFromFile(filename As String) As Long
'***************************************************
'
'
'***************************************************
    CreateTextureFromFile = video_create_texture_from_file(filename)
End Function

Public Function CreateVertexBuffer(ByVal size As Long, ByVal dataPtr As Long, ByVal usage As Long) As Long
'***************************************************
'
'
'***************************************************
    CreateVertexBuffer = video_create_vertex_buffer(size, ByVal dataPtr, usage)
End Function

Public Function CreateIndexBuffer(ByVal size As Long, ByVal dataPtr As Long, ByVal format As Long, ByVal usage As Long) As Long
'***************************************************
'
'
'***************************************************
    CreateIndexBuffer = video_create_index_buffer(size, ByVal dataPtr, format, usage)
End Function

Public Sub SetShaderUniformFloat(ByVal handle As Long, ByVal name As String, ByVal value As Single)
'***************************************************
'
'
'***************************************************
    Call video_set_shader_paramater_float(handle, name, value)
End Sub

Public Sub SetShaderUniformVec2(ByVal handle As Long, ByVal name As String, value As Vec2)
'***************************************************
'
'
'***************************************************
    Call video_set_shader_paramater_float2(handle, name, value)
End Sub

Public Sub SetShaderUniformVec3(ByVal handle As Long, ByVal name As String, value As Vec3)
'***************************************************
'
'
'***************************************************
    Call video_set_shader_paramater_float3(handle, name, value)
End Sub

Public Sub SetShaderUniformVec4(ByVal handle As Long, ByVal name As String, value As Vec4)
'***************************************************
'
'
'***************************************************
    Call video_set_shader_paramater_float4(handle, name, value)
End Sub

Public Sub SetShaderUniformMatrix(ByVal handle As Long, ByVal name As String, value As Matrix)
'***************************************************
'
'
'***************************************************
    Call video_set_shader_paramater_matrix(handle, name, value)
End Sub

Public Sub UpdateVertexBuffer(ByVal handle As Long, ByVal vPtr As Long, ByVal size As Long)
'***************************************************
'
'
'***************************************************
    Call video_update_vertex_buffer(handle, ByVal vPtr, size)
End Sub

Public Sub UpdateIndexBuffer(ByVal handle As Long, ByVal vPtr As Long, ByVal size As Long)
'***************************************************
'
'
'***************************************************
    Call video_update_index_buffer(handle, ByVal vPtr, size)
End Sub

Public Sub SetRenderTarget(ByVal handle As Long)
'***************************************************
'
'
'***************************************************
    Call video_set_render_target(handle)
End Sub

Public Sub BeginRender()
    Call SetRenderTarget(SceneTexture)
    Call ClearBuffer(0)
End Sub

Public Sub EndRender()
    Call SetRenderTarget(-1)
End Sub

Public Sub RenderScene(ByVal deltaTime As Single)
    
End Sub
