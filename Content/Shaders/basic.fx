float4x4 g_view;
/*sampler sam0: register(s0) = 
{
	MinFilter = Linear;
    MagFilter = Linear;
    MipFilter = None;
}*/

sampler sam0: register(s0);

float3 LuminanceConv = { 0.7125f, 0.7154f, 0.0721f };

struct VertexShaderInput
{
    float2 Position : POSITION0;
	float2 TextCoors : TEXCOORD0;
	float4 Color : COLOR0;
};

struct VertexShaderOutput
{
    float4 Position : POSITION0;
	float4 Color   : COLOR0;
	float2 TextCoors : TEXCOORD0;
};

VertexShaderOutput VertexShaderFunction(VertexShaderInput input)
{
    VertexShaderOutput output;
	float4 pos = mul(float4(input.Position,0.0, 1.0), g_view); 
	pos.x += 0.0005;
	pos.y += 0.0005;
    
	output.Color =  input.Color;
	output.TextCoors = input.TextCoors;
    output.Position = pos;
    return output;
}

float4 PixelShaderFunction(VertexShaderOutput input) : COLOR0
{
	return tex2D(sam0, input.TextCoors); 
}

technique Technique1
{
    pass Pass1
    {
        VertexShader = compile vs_2_0 VertexShaderFunction();
        PixelShader = compile ps_2_0 PixelShaderFunction();
    }
}	


