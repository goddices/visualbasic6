// alpha.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "alpha.h"
#include "math.h"

const __int64 MASKR=0xF800F800F800F800;
const __int64 MASKG=0x07E007E007E007E0;
const __int64 MASKB=0x001F001F001F001F;
const __int64 MASKSHIFT= 0x0001000100010001;
const __int64 ADD64 = 0x0040004000400040;
static WORD tablex[1024];
static WORD tabley[1024];

static int costablex[1024];
static int costabley[1024];
static int sintablex[1024];
static int sintabley[1024];

extern _stdcall long bitsmovr(long data,unsigned char bits);
extern _stdcall long bitsmovl(long data,unsigned char bits);
extern _stdcall long getbit(long data,unsigned char bits);
extern _stdcall setpixel(unsigned char *lpDst,long x,long y,long pitch,WORD color);
extern _stdcall unsigned short getpixel(WORD *src,long x,long y,long pitch);
extern _stdcall unsigned short blendcolor(WORD src0,WORD src1,unsigned char alph);
extern _stdcall DrawRect(long x,long y,long w,long h,long spitch, unsigned char alpha, WORD *src);
extern _stdcall DrawAlpha(long x,long y,long sx,long sy,long w,long h,long spitch,long dpitch, unsigned char alpha,WORD keycolor, WORD *dst,WORD *src);
extern _stdcall AddActive(long x,long y,long sx,long sy,long w,long h,long spitch,long dpitch, unsigned char alpha,WORD keycolor,WORD *dst,WORD *src);
extern _stdcall fastmemset(unsigned char *dst,long size,unsigned char data);
extern _stdcall memsetw(WORD *dst,long wordsize,WORD data);
extern _stdcall memcopy_mmx(unsigned char *dst,unsigned char *src,long size);
extern _stdcall memrecopy_mmx(unsigned char *dst,unsigned char *src,long size);
extern _stdcall bltzoom_565_mmx(WORD *lpDst,long iDstX,long iDstY,long iDstPitch,
		long iDstW,long iDstH,WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch,
		long iSrcW,long iSrcH,WORD keycolor);
//缩放特效的支持
extern _stdcall bltzoom_additive_565_mmx(unsigned char alpha,WORD *lpDst,long iDstX,long iDstY,long iDstPitch,
		long iDstW,long iDstH,WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch,
		long iSrcW,long iSrcH,WORD keycolor);
extern _stdcall bltzoom_ablend_565_mmx(unsigned char alpha,WORD *lpDst,long iDstX,long iDstY,long iDstPitch,
		long iDstW,long iDstH,WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch,
		long iSrcW,long iSrcH,WORD keycolor);
//动态光照
extern _stdcall blt_to_lighttable_mmx(unsigned char *lighttable,long iDstX,long iDstY,long w, long h,long iDstPitch,WORD *lpSrc,long iSrcPitch);
//快速Addtive
extern _stdcall fast_additive_565_mmx(unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, 
								 unsigned char *lpDst,long iDstX,long iDstY, long iDstPitch,
								 long iDstW, long iDstH);

extern _stdcall bltfast(char *lpDst,long iDstX,long iDstY,long iDstPitch,
	long iDstW, long iDstH,char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch);
extern _stdcall Qmemset(void *dst, int c, unsigned long nQWORDs);
extern _stdcall Qmemcpy(void *dst, void *src, long nQWORDs);
extern _stdcall bltfast_mmx(WORD *lpDst,long iDstX,long iDstY,long iDstPitch,
	long iDstW, long iDstH,WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch,
	WORD keycolor);
extern _stdcall bltfastex(char *lpdst,long x,long y,long dpitch,long w,
	long h,char *lpsrc,long sx,long sy,long spitch,WORD keycolor);
extern _stdcall ablend_565_mmx(unsigned char alpha,unsigned char *lpSrc,
	long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor);
extern _stdcall colorblend_565(unsigned char alpha, WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, 
	WORD *lpDst,long iDstX,long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor,WORD blendcolor);
extern _stdcall mask_565_mmx(unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, 
	unsigned char *lpDst,long iDstX,long iDstY, long iDstPitch,long iDstW, long iDstH,WORD maskcolor,WORD keycolor);
extern _stdcall halfablend_565_mmx(unsigned char *lpSrc,long iSrcX,
	long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH);
extern _stdcall additive_565_mmx(unsigned char alpha,unsigned char *lpSrc,
	long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor);
extern _stdcall light_565_mmx(unsigned char alpha,unsigned char *lpSrc,
	long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor);
extern _stdcall subitive_565_mmx(unsigned char alpha,unsigned char *lpSrc,
	long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor);
extern _stdcall addcolor_565_mmx(unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, 
	long iDstW, long iDstH,WORD color);
extern _stdcall addlightex_565_mmx(unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor,WORD color);
//Z_Buffer
extern _stdcall zbuffer_blt_mmx(unsigned char *lpscreen,unsigned char *lpzbuffer,long x,long y,WORD z,long scrw,unsigned char *lpsrc,long sx,long sy,long spitch,long sw,long sh,WORD keycolor);

//使用光照表快速处理
extern _stdcall blt_to_lighttable_mmx(unsigned char *lighttable,long iDstX,long iDstY,long iDstW, long iDstH,long iDstPitch,WORD *lpSrc,long iSrcX,long iSrcY,long iSrcPitch);

extern _stdcall alpharect_565_mmx(unsigned char alpha,unsigned char *lpSrc,
	long iSrcX, long iSrcY,long iSrcPitch,long iDstW, long iDstH);
extern _stdcall scanx_565(unsigned char *lpSrc,long iSrcX, long iSrcY,
	long iSrcPitch, unsigned char *lpDst,long iDstX,long iDstY, 
	long iDstPitch,long iDstW, long iDstH,WORD color,WORD keycolor);
extern _stdcall scan_linexy(unsigned char *lpdst,long x, long y,
	long dpitch, unsigned char *lpsrc,long sx,long sy, 
	long spitch,long w, long h,WORD color,WORD keycolor);
//水波
extern _stdcall ripplespread(long *lpbuf,long *lpoldbuf,long w,long h);
extern _stdcall renderipple(WORD *lpscreen,long screenpitch,WORD *lpbmp,long bmppitch,long *lpbuf,long w,long h);
//模糊
extern _stdcall blur_mmx(char *lpscreen,long screenpitch,char *lpbmp,long bmppitch,long w,long h);
extern _stdcall blur_c(WORD *lpscreen,long screenpitch,long x,long y,long w,long h);
//旋转
extern _stdcall rotate_tran(WORD *lpscreen,long screenpitch,WORD *lpbmp,long bmppitch,long x,long y,long dw,long dh,long sx,long sy,long sw,long sh,float angle,WORD keycolor);
//灰度
extern _stdcall gray_565_mmx(char *lpscreen,long screenpitch,char *lpbmp,long bmppitch,long x,long y,long sx,long sy,long w,long h,WORD keycolor);

//Rle Blt
extern _stdcall rle_blt(void *lpdst,long dpitch,long h,long x,long y,void *lpsrc,long pointernum);
extern _stdcall unsigned int RGB565(unsigned int RGB555);
BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}


// This is an example of an exported variable
ALPHA_API int nAlpha=0;

// This is an example of an exported function.
ALPHA_API int fnAlpha(void)
{
	return 42;
}

// This is the constructor of a class that has been exported.
// see alpha.h for the class definition
CAlpha::CAlpha()
{ 
	return; 
}

extern _stdcall long bitsmovl(long data,unsigned char bits)
{
	return data<<bits;
}

extern _stdcall long getbit(long data,unsigned char bits)
{
	//得到指定的二进制位
	return (data>>(bits-1))&1;
}

extern _stdcall long bitsmovr(long data,unsigned char bits)
{
	return data>>bits;
}

extern _stdcall bltfast_mmx(char *lpDst,long iDstX,long iDstY,long iDstPitch,
	long iDstW, long iDstH,char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch,
	WORD keycolor)
{	
	iDstW&=0xfffc;
	if (iDstW<4) return;

	//带透明色不带缩放的16bits Blt
	register __int64 KEY_COLOR=MASKSHIFT*keycolor;
	lpDst+=iDstY*iDstPitch+(iDstX<<1);
	lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);

	iDstPitch-=iDstW<<1;
	iSrcPitch-=iDstW<<1;
	_asm
	{
		mov edi,lpDst
		mov esi,lpSrc
		movq mm7,KEY_COLOR
		mov ebx,iDstH     //loop h
		//shr ebx,1
		ALIGN 8
nextline:
		mov ecx,iDstW
		shr ecx,2
nextpoint:
		movq mm0,[esi]
		movq mm1,mm0
		pcmpeqw mm0,mm7
		psubusw mm1,mm0
		pand mm0,[edi]
		por mm0,mm1
		movq [edi],mm0

		add esi,8
		add edi,8
		dec ecx
		jnz nextpoint

		add edi,iDstPitch
		add esi,iSrcPitch
		dec ebx
		jnz nextline
		emms
	}
}
extern _stdcall bltfastex(char *lpdst,long x,long y,long dpitch,long w,
		long h,char *lpsrc,long sx,long sy,long spitch,WORD keycolor)
{
	//带透明色检查不支持缩放的 16 bits blt
	register __int64 COLORMASK=keycolor*MASKSHIFT;
	lpdst+=y*dpitch+(x<<1);
	lpsrc+=sy*spitch+(sx<<1);
	_asm
	{
		mov edi,lpdst
		mov esi,lpsrc
		mov edx,h
nextline:
		mov ecx,w
		shr ecx,2
nextpoint:
		movq mm0,[esi]
		
		movq [edi],mm0
		add esi,8
		add edi,8
		dec ecx
		jnz nextpoint

		add esi,spitch
		sub esi,w
		sub esi,w
		add edi,dpitch
		sub edi,w
		sub edi,w
		
		dec ecx
		jnz nextline

		emms	//append 2004-11-12
	}
}

extern _stdcall Qmemset(void *dst, int c, unsigned long nQWORDs)
{
	//c的数据？16bits/8bits
	//感谢：<<游戏编程指南>>－彭博
	_asm 
	{
		movq mm0, c
		punpcklbw mm0, mm0
punpcklwd mm0, mm0
punpckldq mm0, mm0
		mov edi, dst

		mov ecx, nQWORDs
		lea	edi, [edi + ecx * 8]
		neg ecx

		movq	mm1, mm0
		movq	mm2, mm0
		movq	mm3, mm0
		movq	mm4, mm0
		movq	mm5, mm0
		movq	mm6, mm0
		movq	mm7, mm0

loopwrite:
		movntq	 [edi + ecx * 8     ], mm0
		movntq	 [edi + ecx * 8 + 8 ], mm1
		movntq	 [edi + ecx * 8 + 16], mm2
		movntq	 [edi + ecx * 8 + 24], mm3
		movntq	 [edi + ecx * 8 + 32], mm4
		movntq	 [edi + ecx * 8 + 40], mm5
		movntq	 [edi + ecx * 8 + 48], mm6
		movntq	 [edi + ecx * 8 + 56], mm7

		add ecx, 8
		jnz loopwrite

		emms
	}
}

extern _stdcall Qmemcpy(void *dst, void *src, long nQWORDs)
{
	//nQWORDs为要拷贝多少个8字节，注意它应能整除8！
	#define CACHEBLOCK 1024 //一个块中有多少QWORDs
								//修改此值有可能实现更高的速度
	int n=((int)(nQWORDs/CACHEBLOCK))*CACHEBLOCK;
	int m=nQWORDs-n;
	if (n)
	{
		_asm //下面先拷贝整数个块
		{
			mov esi, src
			mov edi, dst
			mov ecx, n			//要拷贝多少个块
			lea esi, [esi+ecx*8]
			lea edi, [edi+ecx*8]
			neg ecx
	mainloop:
			mov eax, CACHEBLOCK / 16
	prefetchloop:
			mov ebx, [esi+ecx*8] 		//预读此循环
			mov ebx, [esi+ecx*8+64]	//预读下循环
			add ecx, 16
			dec eax
			jnz prefetchloop
			sub ecx, CACHEBLOCK
			mov eax, CACHEBLOCK / 8
	writeloop:
			movq mm0, qword ptr [esi+ecx*8    ]
			movq mm1, qword ptr [esi+ecx*8+8 ]
			movq mm2, qword ptr [esi+ecx*8+16]
			movq mm3, qword ptr [esi+ecx*8+24]
			movq mm4, qword ptr [esi+ecx*8+32]
			movq mm5, qword ptr [esi+ecx*8+40]
			movq mm6, qword ptr [esi+ecx*8+48]
			movq mm7, qword ptr [esi+ecx*8+56]

			movntq qword ptr [edi+ecx*8   ], mm0
			movntq qword ptr [edi+ecx*8+8 ], mm1
			movntq qword ptr [edi+ecx*8+16], mm2
			movntq qword ptr [edi+ecx*8+24], mm3
			movntq qword ptr [edi+ecx*8+32], mm4
			movntq qword ptr [edi+ecx*8+40], mm5
			movntq qword ptr [edi+ecx*8+48], mm6
			movntq qword ptr [edi+ecx*8+56], mm7
			add ecx, 8
			dec eax
			jnz writeloop
			or ecx, ecx
			jnz mainloop
		}
	}
	if (m)
	{
	_asm
	{
		mov esi, src
		mov edi, dst
		mov ecx, m
		mov ebx, nQWORDs
		lea esi, [esi+ebx*8]
		lea edi, [edi+ebx*8]
		neg ecx
copyloop:
		prefetchnta [esi+ecx*8+512]  //预读
		movq mm0, qword ptr [esi+ecx*8   ]
		movq mm1, qword ptr [esi+ecx*8+8 ]
		movq mm2, qword ptr [esi+ecx*8+16]
		movq mm3, qword ptr [esi+ecx*8+24]
		movq mm4, qword ptr [esi+ecx*8+32]
		movq mm5, qword ptr [esi+ecx*8+40]
		movq mm6, qword ptr [esi+ecx*8+48]
		movq mm7, qword ptr [esi+ecx*8+56]

		movntq qword ptr [edi+ecx*8   ], mm0
		movntq qword ptr [edi+ecx*8+8 ], mm1
		movntq qword ptr [edi+ecx*8+16], mm2
		movntq qword ptr [edi+ecx*8+24], mm3
		movntq qword ptr [edi+ecx*8+32], mm4
		movntq qword ptr [edi+ecx*8+40], mm5
		movntq qword ptr [edi+ecx*8+48], mm6
		movntq qword ptr [edi+ecx*8+56], mm7
		add ecx, 8
		jnz copyloop
		sfence
		emms
	}
}
else
{
	_asm
	{
		sfence
		emms
	}
}

}

extern _stdcall bltfast(WORD *lpDst,long iDstX,long iDstY,long iDstPitch,
	long iDstW, long iDstH,WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch)
{	
	//不带透明色和缩放的16bits Blt
	lpDst+=iDstY*(iDstPitch>>1)+iDstX;
	lpSrc+=iSrcY*(iSrcPitch>>1)+iSrcX;
	iDstW&=0xfffc;
	iDstPitch-=iDstW<<1;
	iSrcPitch-=iDstW<<1;
	_asm
	{
		mov edi,lpDst
		mov esi,lpSrc
		mov edx,iDstW
		mov ebx,iDstH     //loop h
		ALIGN 4
nextline:
		mov ecx,edx
		shr ecx,2
nextpoint:
		movq mm0,[esi]
		movq [edi],mm0
		add esi,8
		add edi,8
		dec ecx
		jnz nextpoint

		add edi,iDstPitch
		add esi,iSrcPitch
		dec ebx
		jnz nextline
		emms
	}
}
extern _stdcall bltzoom_additive_565_mmx(unsigned char alpha,WORD *lpDst,long iDstX,long iDstY,long iDstPitch,
		long iDstW,long iDstH,WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch,
		long iSrcW,long iSrcH,WORD keycolor)
{
	//快速带透明色和缩放的blt,注意：查表不要超过1024
	long i,j,si=0,sj=0,count=0;
	WORD dcol;
	iDstPitch=iDstPitch>>1;
	iSrcPitch=iSrcPitch>>1;
	lpDst+=iDstY*iDstPitch+iDstX;
	lpSrc+=iSrcY*iSrcPitch+iSrcX;
	WORD *tx=tablex;
	WORD *ty=tabley;
	WORD *dline;
	WORD *sline;
	//make tablex
	j=0,count=0;
	for(i=0;i<iDstW;i++)
		for(tablex[i]=2*j,count+=iSrcW;count>=iDstW;j++)
			count-=iDstW;
	//make table[y]
	j=0,count=0;
	for(i=0;i<iDstH;i++)
		for(tabley[i]=j,count+=iSrcH;count>=iDstH;j++)
			count-=iDstH;
	//zoom blt
	sj=-1;
	for(j=0;j<iDstH;j++)
	{
		dline=lpDst+j*iDstPitch;
		sline=lpSrc+tabley[j]*iSrcPitch;
		_asm{

			//asm start
			mov edi,dline
			mov esi,sline
			ALIGN 2;
			mov ecx,iDstW
			mov edx,tx		//edx->表的位置(word)
write:
			mov bx,[edx]	//bx=tablex[i]
			mov eax,esi
			add eax,ebx
			mov bx,[eax]    //src

			cmp bx,keycolor
			jz nextpoint
			
			
			mov ax,[edi]	//处理16比特的象素混合处理 ax->dst,bx->src
		
			movd mm0,eax
			movd mm1,ebx
			movd mm6,alpha

			movq mm2,mm0;//r
			pand mm2,MASKR;
			movq mm3,mm0;//g
			pand mm3,MASKG;
			movq mm4,mm0;//b
			pand mm4,MASKB;


			movq mm5,mm1;//b
			pand mm5,MASKB;
			pmullw mm5,mm6;
			psrlw mm5,8;
			paddusw mm4,mm5;  //  if mm4>MASKB then mm4=MASKB
			movq mm5,mm4;
			pcmpgtw mm5,MASKB;
			por mm4,mm5;
			pand mm4,MASKB;

			movq mm5,mm1;//g
			pand mm5,MASKG;
			psrlw mm5,5;
			pmullw mm5,mm6;
			psrlw mm5,8;
			psllw mm5,5;
			paddusw mm3,mm5;  //  if mm3>MASKG then mm3=MASKG
			movq mm5,mm3;
			pcmpgtw mm5,MASKG;
			por mm3,mm5;
			pand mm3,MASKG;
			por mm4,mm3;

			movq mm5,mm1;//r
			pand mm5,MASKR;
			psrlw mm5,11;
			pmullw mm5,mm6;
			psrlw mm5,8;
			psllw mm5,11;
			paddusw mm2,mm5;  //  if mm2>MASKR then mm2=MASKR
			pand mm2,MASKR;
			por mm4,mm2;


			movd eax,mm4
			mov dcol,ax
			mov [edi],ax

nextpoint:
			add edx,2		//edx=edx+2->2bytes
			add edi,2

			dec ecx
			jnz write
			emms
		}
	}
	_asm
	{
		emms
	}
}

extern _stdcall bltzoom_ablend_565_mmx(unsigned char alpha,WORD *lpDst,long iDstX,long iDstY,long iDstPitch,
		long iDstW,long iDstH,WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch,
		long iSrcW,long iSrcH,WORD keycolor)
{
	//缩放透明
long i,j,si=0,sj=0,count=0;
	unsigned int tmp;
	iDstPitch=iDstPitch>>1;
	iSrcPitch=iSrcPitch>>1;
	lpDst+=iDstY*iDstPitch+iDstX;
	lpSrc+=iSrcY*iSrcPitch+iSrcX;
	
	WORD *tx=tablex;
	WORD *ty=tabley;
	WORD *dline;
	WORD *sline;
	//make tablex
	j=0,count=0;
	for(i=0;i<iDstW;i++)
		for(tablex[i]=2*j,count+=iSrcW;count>=iDstW;j++)
			count-=iDstW;
	//make table[y]
	j=0,count=0;
	for(i=0;i<iDstH;i++)
		for(tabley[i]=j,count+=iSrcH;count>=iDstH;j++)
			count-=iDstH;
	//zoom blt
	sj=-1;
	for(j=0;j<iDstH;j++)
	{
		dline=lpDst+j*iDstPitch;
		sline=lpSrc+tabley[j]*iSrcPitch;
		_asm{

			//asm start
			mov edi,dline
			mov esi,sline
			ALIGN 2;
			mov ecx,iDstW
			mov edx,tx		//edx->表的位置(word)
write:
			mov bx,[edx]	//bx=tablex[i]
			mov eax,esi
			add eax,ebx
			mov bx,[eax]    //src

			cmp bx,keycolor
			jz nextpoint

			mov ax,[edi]	//处理16比特的象素混合处理 ax->dst,bx->src

			movd mm1,eax
			movd mm0,ebx
			movd mm6,alpha

			movq mm2,mm0//r
			movq mm3,mm0//g
			movq mm4,mm0//b
			pand mm2,MASKR
			pand mm3,MASKG
			pand mm4,MASKB

			movq mm5,mm1//b
			pand mm5,MASKB
			psubw mm5,mm4
			pmullw mm5,mm6
			psrlw mm5,8
			paddsw mm4,mm5
			pand mm4,MASKB

			movq mm5,mm1//g
			pand mm5,MASKG
			psrlw mm5,5
			psrlw mm3,5
			psubw mm5,mm3
			pmullw mm5,mm6
			psrlw mm5,8
			paddsw mm3,mm5
			psllw mm3,5
			pand mm3,MASKG
			por mm4,mm3

			movq mm5,mm1//r
			pand mm5,MASKR
			psrlw mm5,11
			psrlw mm2,11
			psubw mm5,mm2
			pmullw mm5,mm6
			psrlw mm5,8
			paddsw mm2,mm5
			psllw mm2,11
			pand mm2,MASKR
			por mm4,mm2
			
			movd eax,mm4			//???怎么回事啊？vb编译后执行就出错！VB的问题还是VC？
//*/
			mov [edi],ax
nextpoint:
			add edx,2		//edx=edx+2->2bytes
			add edi,2

			dec ecx
			jnz write
		}
	}
	_asm
	{
		emms
	}

}

extern _stdcall bltzoom_565_mmx(WORD *lpDst,long iDstX,long iDstY,long iDstPitch,
		long iDstW,long iDstH,WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch,
		long iSrcW,long iSrcH,WORD keycolor)
{
	//快速带透明色和缩放的blt,注意：查表不要超过1024
	long i,j,si=0,sj=0,count=0;
	iDstPitch=iDstPitch>>1;
	iSrcPitch=iSrcPitch>>1;
	lpDst+=iDstY*iDstPitch+iDstX;
	lpSrc+=iSrcY*iSrcPitch+iSrcX;
	WORD *tx=tablex;
	WORD *ty=tabley;
	WORD *dline;
	WORD *sline;
	//make tablex
	j=0,count=0;
	for(i=0;i<iDstW;i++)
		for(tablex[i]=2*j,count+=iSrcW;count>=iDstW;j++)
			count-=iDstW;
	//make table[y]
	j=0,count=0;
	for(i=0;i<iDstH;i++)
		for(tabley[i]=j,count+=iSrcH;count>=iDstH;j++)
			count-=iDstH;
	//zoom blt
	sj=-1;
	for(j=0;j<iDstH;j++)
	{
		dline=lpDst+j*iDstPitch;
		sline=lpSrc+tabley[j]*iSrcPitch;
		_asm{

			//asm start
			mov edi,dline
			mov esi,sline
			ALIGN 2;
			mov ecx,iDstW
			mov edx,tx		//edx->表的位置(word)
write:
			mov bx,[edx]	//bx=tablex[i]
			mov eax,esi
			add eax,ebx
			mov bx,[eax]    //src

			cmp bx,keycolor
			jz nextpoint
			mov [edi],bx
nextpoint:
			add edx,2		//edx=edx+2->2bytes
			add edi,2

			dec ecx
			jnz write
		}
	}
	_asm
	{
		emms
	}

}

extern _stdcall setpixel(unsigned char *lpDst,long x,long y,long pitch,WORD color)
{
	lpDst+=y*pitch+x*2;
	_asm{
		mov edi,lpDst
		mov ax,color
		mov [edi],ax
	}
}

extern _stdcall unsigned short getpixel(WORD *src,long x,long y,long pitch)
{
	//得到16Bits的象素值
	src+=y*pitch/2+x;
	return *src;
}
extern _stdcall unsigned short blendcolor(WORD src0,WORD src1,unsigned char alph)
{
	WORD r0,g0,b0,r1,g1,b1;
	r1=src1 & 0xf800;
	r0=src0 & 0xf800;
	g1=src1 & 0x7e0;
	g0=src0 & 0x7e0;
	b1=src1 & 0x1f;
	b0=src0 & 0x1f;
	r0+=(r1-r0)*alph>>8;
	g0+=(g1-g0)*alph>>8;
	b0+=(b1-b0)*alph>>8;
	return ((b0 & 0x1f)|(g0 & 0x7e0)|(r0 & 0xf800));
}

extern _stdcall blt(unsigned char *lpDst,long iDstX,long iDstY,long iDstPitch,
	long iDstW, long iDstH,unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch)
{	
	//带透明色，缩放的16bits Blt

}

extern _stdcall DrawRect(long x,long y,long w,long h,long spitch, unsigned char alpha, WORD *src)
{
	//alpha 透明效果
	register  WORD r,g,b,color;
	register int i,j;
	register long offset1=y*spitch+x;
	//back
	for(j=0;j<h;j++)		
	{
		for(i=0;i<w;i++)
		{
			color=*(src+offset1);
			r=(0xf800 & color)*alpha>>8;
			g=(0x07e0 & color)*alpha>>8;
			b=(0x001f & color)*alpha>>8;

			color=WORD((b&0x001f)|(g&0x07e0)|(r&0xf800));
			*(src+offset1) =color;

			offset1++;
		}
		offset1+=spitch-w;
	}
}

extern _stdcall AddActive(long x,long y,long sx,long sy,long w,long h,
						  long spitch,long dpitch, unsigned char alpha,
						  WORD keycolor,WORD *dst,WORD *src)
{
	//alpha 透明效果
	register DWORD r,g,b,r0,g0,b0;
	register WORD color,color1;
	register long offset1=y*dpitch+x,offset2=sy*spitch+sx;
	register int i,j;
	for(j=0;j<h;j++)		
	{
		for(i=0;i<w;i++)
		{
			color1=*(src+offset2);
			if (keycolor==color1) 
			{
				offset1++;
				offset2++;
				continue;
			}
			color=*(dst+offset1);
			r=0xf800 & color;
			g=0x07e0 & color;
			b=0x001f & color;

			r0=0xf800 & color1;
			g0=0x07e0 & color1;
			b0=0x001f & color1;
			
			r+=r0*alpha>>8;
			g+=g0*alpha>>8;
			b+=b0*alpha>>8;
			if (r>0xf800) r=0xf800;
			if (g>0x07e0) g=0x07e0;
			if (b>0x001f) b=0x001f;
			
			color=WORD((b&0x001f)|(g&0x07e0)|(r&0xf800));
			*(dst+offset1) =color;
			
			offset1++;
			offset2++;
		}
		offset1+=dpitch-w;
		offset2+=spitch-w;
	}
}
extern _stdcall DrawAlpha(long x,long y,long sx,long sy,long w,long h,long spitch,long dpitch, unsigned char alpha,WORD keycolor,WORD *dst,WORD *src)
{
	//alpha 透明效果
	register WORD r,g,b,r0,g0,b0;
	register WORD color,color1;
	register long offset1=y*dpitch+x,offset2=sy*spitch+sx;
	register int i,j;
	for(j=0;j<h;j++)		
	{
		for(i=0;i<w;i++)
		{
			color1=*(src+offset2);
			if (keycolor==color1) 
			{
				offset1++;
				offset2++;
				continue;
			}
			color=*(dst+offset1);
			r=0xf800 & color;
			g=0x07e0 & color;
			b=0x001f & color;

			r0=0xf800 & color1;
			g0=0x07e0 & color1;
			b0=0x001f & color1;
			
			r+=(r0-r)*alpha>>8;
			g+=(g0-g)*alpha>>8;
			b+=(b0-b)*alpha>>8;
			
			color=WORD((b&0x001f)|(g&0x07e0)|(r&0xf800));
			*(dst+offset1) =color;
			
			offset1++;
			offset2++;
		}
		offset1+=dpitch-w;
		offset2+=spitch-w;
	}
}


//MMX加速指令
extern _stdcall memcopy_mmx(unsigned char *dst,unsigned char *src,long size)
{
	//内存COPY size bytes
	long BYTES=size>>3;
	_asm
	{
		mov esi,src
		mov edi,dst
		mov ecx,BYTES
NEXT:
		movq mm0,[esi]
		movq [edi],mm0
		add esi,8
		add edi,8
		dec ecx
		jnz NEXT
		emms
	}
}
extern _stdcall memrecopy_mmx(unsigned char *dst,unsigned char *src,long size)
{
	//内存的反向Copy
	long BYTES=size>>3;
	_asm
	{
		mov esi,src
		mov edi,dst
		mov ecx,BYTES
NEXT:
		movq mm0,[esi]
		movq [edi],mm0
		sub esi,8
		sub edi,8
		dec ecx
		jnz NEXT
		emms
	}

}

extern _stdcall fastmemset(unsigned char *dst,long bytesize,unsigned char data)
{
	_asm
	{
		//
		
		mov ecx,bytesize
		mov edi,dst
		mov al,data
		rep stosb
	}

}

extern _stdcall memsetw(WORD *dst,long wordsize,WORD data)
{
	_asm
	{
		//
		
		mov ecx,wordsize
		mov edi,dst
		mov ax,data
		rep stosw
	}

}

extern _stdcall memset_mmx(unsigned char *dst,long pitch,long w,long h,unsigned char data)
{
	//
	register __int64 data64=0x0101010101010101*data;
	pitch-=2*w;
	_asm
	{
		//asm start
		movq mm0,data64
		mov edi,dst
		mov edx,h
nextline:
		mov ecx,w
		shr ecx,2
nextpoint:
		movq [edi],mm0
		//mov [edi],255
		add edi,8
		dec ecx
		jnz nextpoint

		add edi,pitch

		dec edx
		jnz nextline
		emms
	}
}
extern _stdcall halfablend_565_mmx(unsigned char *lpSrc,long iSrcX,
	long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH)
{

	register __int64 MASK=0x7bef7bef7bef7bef;

	lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);
	lpDst+=iDstY*iDstPitch+(iDstX<<1);
	
	register long w=iDstW<<1;
	_asm
	{
		mov eax,iSrcPitch
		sub eax,w
		mov ebx,iDstPitch
		sub ebx,w

		mov esi,lpSrc;
		mov edi,lpDst;
		mov edx,iDstH;//H
START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		
		movq mm0,[esi];		//Get Data
		movq mm1,[edi];
		psrlq mm0,1;
		pand mm0,MASK;
		psrlq mm1,1;
		pand mm1,MASK;
		paddusw mm0,mm1;

		movq [esi],mm0
		add esi,8
		add edi,8

		dec ecx
		jnz NEXT_POINT
//NEXT_LINE:
		add esi,eax
		add edi,ebx

		dec edx;
		jnz START;
//DONE:
		EMMS;
	}
}

extern _stdcall light_565_mmx(unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor)
{
	//表现光线的效果

	register __int64 ALPHA_MASK;
				//ALPHA_KEYCOLOR=MASKSHIFT*keycolor;
	register __int64 ALPHA_KEYMASK=0xffffffffffffffff;
	register __int64 ALPHA_KEYCOLOR=keycolor*MASKSHIFT;
	
	lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);
	lpDst+=iDstY*iDstPitch+(iDstX<<1);

	register long w=iDstW<<1;
	_asm
	{
		mov eax,iSrcPitch
		sub eax,w
		mov ebx,iDstPitch
		sub ebx,w

		mov esi,lpSrc;
		mov edi,lpDst;
		mov edx,iDstH;//H
START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		
		movq mm0,[esi];		//Get Data
		movq mm1,[edi];

		movq mm5,mm1;
        pand mm5,MASKG;
		psrlw mm5,3;
		
		movq mm6,mm5;
		movq mm5,mm1;
		pcmpeqw mm5,ALPHA_KEYCOLOR;
		pxor mm5,ALPHA_KEYMASK;
		pand mm6,mm5;
		pxor mm7,mm7;

		movq mm2,mm0;//r
		pand mm2,MASKR;
		movq mm3,mm0;//g
		pand mm3,MASKG;
		movq mm4,mm0;//b
		pand mm4,MASKB;


		movq mm5,mm1;//b
		pand mm5,MASKB;
		pmullw mm5,mm6;
		psrlw mm5,8;
		paddusw mm4,mm5;  //  if mm4>MASKB then mm4=MASKB
		movq mm5,mm4;
		pcmpgtw mm5,MASKB;
		por mm4,mm5;
		pand mm4,MASKB;
		paddusw mm7,mm4;

		movq mm5,mm1;//g
		pand mm5,MASKG;
		psrlw mm5,5;
		pmullw mm5,mm6;
		psrlw mm5,8;
		psllw mm5,5;
		paddusw mm3,mm5;  //  if mm3>MASKG then mm3=MASKG
		movq mm5,mm3;
		pcmpgtw mm5,MASKG;
		por mm3,mm5;
		pand mm3,MASKG;
		paddusw mm7,mm3;

		movq mm5,mm1;//r
		pand mm5,MASKR;
		psrlw mm5,11;
		pmullw mm5,mm6;
		psrlw mm5,8;
		psllw mm5,11;
		paddusw mm2,mm5;  //  if mm2>MASKR then mm2=MASKR
		pand mm2,MASKR;
		paddusw mm7,mm2;


		movq [esi],mm7;
		add esi,8;
		add edi,8;
		dec ecx;
		jnz NEXT_POINT;

//NEXT_LINE:
		add esi,eax;
		add edi,ebx;

		dec edx;
		jnz START;
//DONE:
		EMMS;
	}
}
extern _stdcall fastlight_565_mmx(unsigned char *lpDst,unsigned char *lpSrc,unsigned char *lpTable,
	long iDstX,long iDstY,long iDstPitch,long iSrcPitch,long iTablePitch,long iDstW,long iDstH)
{
	//光照表的实现
	lpDst+=iDstY*iDstPitch+(iDstX<<1);

	register long w=iDstW<<1;
	_asm
	{
		//mov eax,iSrcPitch;
		mov eax,lpTable;	//eax 光照表的地址
		mov ebx,iDstPitch;
		mov esi,lpSrc;
		mov edi,lpDst;
		ALIGN 8;

		mov edx,iDstH;//H
START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		
		movq mm0,[esi];		//Get Data
		movd mm1,[eax];

		pxor mm7,mm7;
		punpcklbw mm1,mm7;

		movq mm2,mm0   //b
		movq mm3,mm0   //g
		pand mm2,MASKB
		movq mm4,mm0   //r
		pand mm3,MASKG
		pand mm4,MASKR

		pmullw mm2,mm1;	//b
		psrlw mm2,8;
//		pand mm2,MASKB;

		psrlw mm3,5;
		pmullw mm3,mm1;
		psrlw mm3,8;
		psllw mm3,5;
//		pand mm3,MASKG;

		psrlw mm4,11;
		pmullw mm4,mm1;
		psrlw mm4,8;
		psllw mm4,11;
//		pand mm4,MASKR;

		paddusw mm2,mm3;
		paddusw mm2,mm4;
		//psllw
		//paddusw mm7,mm3;

		movq [edi],mm2;
		add esi,8;
		add edi,8;
		add eax,4;

		dec ecx
		jnz NEXT_POINT
//NEXT_LINE:
		add esi,iSrcPitch
		sub esi,w

		add edi,ebx
		sub edi,w

		add eax,iTablePitch
		sub eax,iDstW

		dec edx
		jnz START
//DONE:
		EMMS;
	}
}
extern _stdcall addlightex_565_mmx(unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor,WORD color)
{
	//表现光线的效果

	register __int64 ALPHA_KEYMASK=0xffffffffffffffff;
	register __int64 ALPHA_KEYCOLOR=keycolor*MASKSHIFT;
	register __int64 ADD_COLOR=MASKSHIFT*color;	

	lpDst+=iDstY*iDstPitch+(iDstX<<1);
	lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);

	register long w=iDstW<<1;
	_asm
	{
		mov eax,iSrcPitch;
		sub eax,w
		mov ebx,iDstPitch;
		sub ebx,w

		mov esi,lpSrc;
		mov edi,lpDst;
		mov edx,iDstH;//H

START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		
		movq mm0,[esi];		//Get Data
		movq mm1,[edi];

		movq mm6,ADD_COLOR;
		movq mm5,mm1;
		pcmpeqw mm5,ALPHA_KEYCOLOR;
		pand mm0,mm5;
		pxor mm5,ALPHA_KEYMASK;
		pand mm6,mm5;
		pand mm1,mm5;
		por mm1,mm0;
		pxor mm7,mm7;

		movq mm2,mm1;//r
		pand mm2,MASKR;
		movq mm3,mm1;//g
		pand mm3,MASKG;
		movq mm4,mm1;//b
		pand mm4,MASKB;

		movq mm5,mm6;//b
		pand mm5,MASKB;
		paddusw mm4,mm5;  //  if mm4>MASKB then mm4=MASKB
		movq mm5,mm4;
		pcmpgtw mm5,MASKB;
		por mm4,mm5;
		pand mm4,MASKB;
		paddusw mm7,mm4;

		movq mm5,mm6;//g
		pand mm5,MASKG;
		paddusw mm3,mm5;  //  if mm3>MASKG then mm3=MASKG
		movq mm5,mm3;
		pcmpgtw mm5,MASKG;
		por mm3,mm5;
		pand mm3,MASKG;
		paddusw mm7,mm3;

		movq mm5,mm6;//r
		pand mm5,MASKR;
		paddusw mm2,mm5;  //  if mm2>MASKR then mm2=MASKR
		pand mm2,MASKR;
		paddusw mm7,mm2;

		movq [esi],mm7;
		add esi,8;
		add edi,8;

		dec ecx;
		jnz NEXT_POINT;
//NEXT_LINE:
		add esi,eax;
		add edi,ebx;

		dec edx;
		jnz START;
//DONE:
		EMMS;
	}
}

extern _stdcall addcolor_565_mmx(unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, 
	long iDstW, long iDstH,WORD color)
{
	//图象加色
	long SrcOffset=iSrcY*iSrcPitch+iSrcX*2;
	register __int64 ADD_COLOR=MASKSHIFT*color;	
	register long w=2*iDstW;
	_asm
	{
		mov eax,iSrcPitch;
		mov esi,lpSrc;
		add esi,SrcOffset;
		mov edx,iDstH;//H
START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		
		movq mm0,[esi];		//Get Data
 
		movq mm2,mm0;//r
		pand mm2,MASKR;
		movq mm3,mm0;//g
		pand mm3,MASKG;
		movq mm4,mm0;//b
		pand mm4,MASKB;
		pxor mm7,mm7;

		movq mm5,ADD_COLOR;//b
		pand mm5,MASKB;
		paddusw mm4,mm5;  //  if mm4>MASKB then mm4=MASKB
		movq mm5,mm4;
		pcmpgtw mm5,MASKB;
		por mm4,mm5;
		pand mm4,MASKB;
		paddusw mm7,mm4;

		movq mm5,ADD_COLOR;//g
		pand mm5,MASKG;
		paddusw mm3,mm5;  //  if mm3>MASKG then mm3=MASKG
		movq mm5,mm3;
		pcmpgtw mm5,MASKG;
		por mm3,mm5;
		pand mm3,MASKG;
		paddusw mm7,mm3;

		movq mm5,ADD_COLOR;//r
		pand mm5,MASKR;
		paddusw mm2,mm5;  //  if mm2>MASKR then mm2=MASKR
		pand mm2,MASKR;
		paddusw mm7,mm2;

		movq [esi],mm7;
		add esi,8;
		dec ecx;
		cmp ecx,0;
		jnz NEXT_POINT;
//NEXT_LINE:
		add esi,eax;
		sub esi,w;

		dec edx;
		cmp edx,0;
		jnz START;
//DONE:
		EMMS;
	}
}

extern _stdcall fast_additive_565_mmx(unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, 
								 unsigned char *lpDst,long iDstX,long iDstY, long iDstPitch,
								 long iDstW, long iDstH)
{
	//不带透明，alpha参数的Addtive
	lpDst+=iDstY*iDstPitch+(iDstX<<1);
	lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);

	register long w=iDstW<<1;
	_asm
	{
		mov eax,iSrcPitch
		sub eax,w
		mov ebx,iDstPitch
		sub ebx,w

		mov esi,lpSrc
		mov edi,lpDst

		mov edx,iDstH//H
START:
		mov ecx,iDstW// iDstW/4
		shr ecx,2

NEXT_POINT:
		//Thanks For 云风 
		movq mm0,[esi]
		movq mm3,[edi]
		movq mm1,mm0
		movq mm2,mm0
		movq mm4,mm3
		movq mm5,mm3
		psllw mm1,5
		psllw mm4,5
		psllw mm2,11
		psllw mm5,11
		paddusw mm0,mm3 // 红色加
		paddusw mm1,mm4 // 绿色加
		paddusw mm2,mm5 // 蓝色加
		psrlw mm1,5 
		psrlw mm2,11 
		pand mm0,MASKR
		pand mm1,MASKG
		por mm0,mm2
		por mm0,mm1
		movq [edi],mm0

		add esi,8
		add edi,8

		dec ecx
		jnz NEXT_POINT
//NEXT_LINE:
		add esi,eax
		add edi,ebx

		dec edx
		jnz START
//DONE:
		emms
	}
}
extern _stdcall additive_565_mmx(unsigned char alpha,unsigned char *lpSrc,
	long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor)
{

	register __int64 ALPHA_MASK=MASKSHIFT*alpha;
				//ALPHA_KEYCOLOR=MASKSHIFT*keycolor;
	register __int64 ALPHA_KEYMASK=0xffffffffffffffff;
	register __int64 ALPHA_KEYCOLOR=keycolor*MASKSHIFT;

	lpDst+=iDstY*iDstPitch+(iDstX<<1);
	lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);

	register long w=iDstW<<1;
	_asm
	{
		mov eax,iSrcPitch;
		sub eax,w
		mov ebx,iDstPitch;
		sub ebx,w

		mov esi,lpSrc;
		mov edi,lpDst;

		mov edx,iDstH;//H
START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		
		movq mm0,[esi];		//Get Data
		movq mm1,[edi];

		movq mm6,ALPHA_MASK;
		movq mm5,mm1;
		pcmpeqw mm5,ALPHA_KEYCOLOR;
		pxor mm5,ALPHA_KEYMASK;
		pand mm6,mm5;
		pxor mm7,mm7;

		movq mm2,mm0;//r
		pand mm2,MASKR;
		movq mm3,mm0;//g
		pand mm3,MASKG;
		movq mm4,mm0;//b
		pand mm4,MASKB;


		movq mm5,mm1;//b
		pand mm5,MASKB;
		pmullw mm5,mm6;
		psrlw mm5,8;
		paddusw mm4,mm5;  //  if mm4>MASKB then mm4=MASKB
		movq mm5,mm4;
		pcmpgtw mm5,MASKB;
		por mm4,mm5;
		pand mm4,MASKB;
		paddusw mm7,mm4;

		movq mm5,mm1;//g
		pand mm5,MASKG;
		psrlw mm5,5;
		pmullw mm5,mm6;
		psrlw mm5,8;
		psllw mm5,5;
		paddusw mm3,mm5;  //  if mm3>MASKG then mm3=MASKG
		movq mm5,mm3;
		pcmpgtw mm5,MASKG;
		por mm3,mm5;
		pand mm3,MASKG;
		paddusw mm7,mm3;

		movq mm5,mm1;//r
		pand mm5,MASKR;
		psrlw mm5,11;
		pmullw mm5,mm6;
		psrlw mm5,8;
		psllw mm5,11;
		paddusw mm2,mm5;  //  if mm2>MASKR then mm2=MASKR
		pand mm2,MASKR;
		paddusw mm7,mm2;


		movq [esi],mm7;
		add esi,8;
		add edi,8;

		dec ecx;
		jnz NEXT_POINT;
//NEXT_LINE:
		add esi,eax;
		add edi,ebx;

		dec edx;
		jnz START;
//DONE:
		EMMS;
	}
}

extern _stdcall alpharect_565_mmx(unsigned char alpha,unsigned char *lpSrc,
	long iSrcX, long iSrcY,long iSrcPitch,long iDstW, long iDstH)
{
	long SrcOffset=iSrcY*iSrcPitch+iSrcX*2;
	register __int64 ALPHA_MASK=MASKSHIFT*alpha;
				//ALPHA_KEYCOLOR=MASKSHIFT*keycolor;
	register __int64 ALPHA_KEYMASK=0xffffffffffffffff;
	//keycolor=0;

	register long w=2*iDstW;
	_asm
	{
		mov eax,iSrcPitch;

		mov esi,lpSrc;
		add esi,SrcOffset;

		mov edx,iDstH;//H
START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		
		movq mm0,[esi];		//Get Data

		movq mm6,ALPHA_MASK;
		pxor mm7,mm7;

		movq mm5,mm0;//b
		pand mm5,MASKB;
		pmullw mm5,mm6;
		psrlw mm5,8;
		pand mm5,MASKB;
		paddusw mm7,mm5;

		movq mm5,mm0;//g
		pand mm5,MASKG;
		psrlw mm5,5;
		pmullw mm5,mm6;
		psrlw mm5,8;
		psllw mm5,5;
		pand mm5,MASKG;
		paddusw mm7,mm5;

		movq mm5,mm0;//r
		pand mm5,MASKR;
		psrlw mm5,11;
		pmullw mm5,mm6;
		psrlw mm5,8;
		psllw mm5,11;
		pand mm5,MASKR;
		paddusw mm7,mm5;

		movq [esi],mm7;
		add esi,8;
		dec ecx;
		cmp ecx,0;
		jnz NEXT_POINT;
//NEXT_LINE:
		add esi,eax;
		sub esi,w;

		dec edx;
		cmp edx,0;
		jnz START;
//DONE:
		EMMS;
	}
}

extern _stdcall subitive_565_mmx(unsigned char alpha,unsigned char *lpSrc,
	long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor)
{
	lpDst+=iDstY*iDstPitch+(iDstX<<1);
	lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);

	register __int64 ALPHA_MASK=MASKSHIFT*alpha;
				//ALPHA_KEYCOLOR=MASKSHIFT*keycolor;
	register __int64 ALPHA_KEYMASK=0xffffffffffffffff;
	register __int64 ALPHA_KEYCOLOR=keycolor*MASKSHIFT;

	register long w=iDstW<<1;
	_asm
	{
		mov eax,iSrcPitch;
		sub eax,w
		mov ebx,iDstPitch;
		sub ebx,w

		mov esi,lpSrc;
		mov edi,lpDst;

		mov edx,iDstH;//H
START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		
		movq mm0,[esi];		//Get Data
		movq mm1,[edi];

		movq mm6,ALPHA_MASK;
		movq mm5,mm1;
		pcmpeqw mm5,ALPHA_KEYCOLOR;
		pxor mm5,ALPHA_KEYMASK;
		pand mm6,mm5;
		pxor mm7,mm7;

		movq mm2,mm0;//r
		pand mm2,MASKR;
		movq mm3,mm0;//g
		pand mm3,MASKG;
		movq mm4,mm0;//b
		pand mm4,MASKB;


		movq mm5,mm1;//b
		pand mm5,MASKB;
		pmullw mm5,mm6;
		psrlw mm5,8;
		psubusw mm4,mm5;  
		pand mm4,MASKB;
		paddusw mm7,mm4;

		movq mm5,mm1;//g
		pand mm5,MASKG;
		psrlw mm5,5;
		pmullw mm5,mm6;
		psrlw mm5,8;
		psrlw mm3,5;
		psubusw mm3,mm5;  
		psllw mm3,5;
		pand mm3,MASKG;
		paddusw mm7,mm3;

		movq mm5,mm1;//r
		pand mm5,MASKR;
		psrlw mm5,11;
		pmullw mm5,mm6;
		psrlw mm5,8;
		psrlw mm2,11;
		psubusw mm2,mm5;  //  if mm2>MASKR then mm2=MASKR
		psllw mm2,11;
		pand mm2,MASKR;
		paddusw mm7,mm2;

		movq [esi],mm7

		add esi,8;
		add edi,8;

		dec ecx
		jnz NEXT_POINT;
//NEXT_LINE:
		add esi,eax;
		add edi,ebx;

		dec edx
		jnz START;
//DONE:
		EMMS;
	}
}
extern _stdcall colorblend_565(unsigned char alpha, WORD *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, 
	WORD *lpDst,long iDstX,long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor,WORD blendcolor)
{
	//图象加色处理
	iSrcPitch=iSrcPitch>>1;
	iDstPitch=iDstPitch>>1;
	long SrcOffset=iSrcY*iSrcPitch+iSrcX,
				  DstOffset=iDstY*iDstPitch+iDstX;
	unsigned char r0,g0,b0,r1,g1,b1;
	WORD color;
	int i=iDstW,j=iDstH;
	lpSrc+=SrcOffset;
	lpDst+=DstOffset;
	r0=(blendcolor>>11)&0x1f;
	g0=(blendcolor>>5)&0x3f;
	b0=blendcolor&0x1f;
	for(j=0;j<iDstH;j++)
	{
		for(i=0;i<iDstW;i++)
		{

			if(*lpDst!=keycolor)
			{
				color=*lpDst;
				r1=(color>>11)&0x1f;
				g1=(color>>5)&0x3f;
				b1=color&0x1f;	
				r1+=r0;
				g1+=g0;
				b1+=b0;
				if(r1>0x1f) r1=0x1f;
				if(g1>0x3f) g1=0x3f;
				if(b1>0x1f) b1=0x1f;
				color=(r1<<11)|(g1<<5)|(b1);
				*lpSrc=color;
			}
			lpSrc++;
			lpDst++;
		}
		lpSrc+=iSrcPitch-iDstW;
		lpDst+=iDstPitch-iDstW;
	}
}
extern _stdcall mask_565_mmx(unsigned char *lpSrc,long iSrcX, long iSrcY,long iSrcPitch, 
	unsigned char *lpDst,long iDstX,long iDstY, long iDstPitch,long iDstW, long iDstH,
	WORD maskcolor,WORD keycolor)
{
	//green Mask opretion
	lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);
	lpDst+=iDstY*iDstPitch+(iDstX<<1);

	register __int64 ALPHA_KEYMASK=0xffffffffffffffff;
	register __int64 ALPHA_KEYCOLOR=keycolor*MASKSHIFT;
	register __int64 COLOR_MASK=maskcolor*MASKSHIFT;
	register long w=iDstW<<1;
	_asm
	{
		mov eax,iSrcPitch
		sub eax,w

		mov ebx,iDstPitch
		sub ebx,w

		mov esi,lpSrc
		mov edi,lpDst

		mov edx,iDstH;//H
START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		movq mm0,[esi];
		movq mm1,[edi];     //Get Data

		movq mm6,COLOR_MASK;
		movq mm5,mm1;
		pcmpeqw mm5,ALPHA_KEYCOLOR;
		pand mm0,mm5;       //mm0 por mm1-
		pxor mm5,ALPHA_KEYMASK;
		pand mm6,mm5;
		pand mm1,mm6;
		por mm1,mm0;
		movq [esi],mm1;

		add esi,8;
		add edi,8;
		dec ecx;

		jnz NEXT_POINT;
//NEXT_LINE:
		add esi,eax
		add edi,ebx

		dec edx
		jnz START
//DONE:
		EMMS;
	}
}
extern _stdcall scanx_565(unsigned char *lpSrc,long iSrcX, long iSrcY,
	long iSrcPitch, unsigned char *lpDst,long iDstX,long iDstY, 
	long iDstPitch,long iDstW, long iDstH,WORD color,WORD keycolor)
{
	//提取轮廓线
	register long w=iDstW<<1;

	register __int32 f=0xffffffff;
	register __int32 KEYCOLOR_32=keycolor*0x10001;
	register __int32 spitch=iSrcPitch-w;
	register __int32 dpitch=iDstPitch-w;

    lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);
	lpDst+=iDstY*iDstPitch+(iDstX<<1);

	_asm
	{
		mov esi,lpSrc
		mov edi,lpDst

		mov edx,iDstH;//H
START:
		mov ecx,iDstW// 32bits
		shr ecx,1
NEXT_POINT:
		mov eax,[esi];
		mov ebx,eax;
		movd mm0,ebx;	

		pcmpeqw mm0,KEYCOLOR_32
		movd ebx,mm0
		xor ebx,f

		cmp ebx,0
		jz MOVE_NEXT

		not f
		mov bx,color	//水平轮廓线
		mov [edi],bx

MOVE_NEXT:
		add esi,4
		add edi,4

		dec ecx
		jnz NEXT_POINT
//NEXT_LINE:
		add esi,spitch
		add edi,dpitch

		dec edx
		jnz START
//DONE:
		emms
	}
}

extern _stdcall scan_linexy(WORD *lpdst,long x, long y,
	long dpitch, WORD *lpsrc,long sx,long sy, 
	long spitch,long w, long h,WORD color,WORD keycolor)
{
	//扫描得到轮廓线
	spitch=spitch>>1;
	dpitch=dpitch>>1;

	WORD *dst,*src;
	lpdst+=dpitch*y+x;
	lpsrc+=spitch*sy+sx;

	int i,j;
	for(i=1;i<h-1;i++)
	{
		dst=lpdst+dpitch*i+1;
		src=lpsrc+spitch*i+1;
		for(j=1;j<w-1;j++)
		{
			//if (i,j)透明and 有不透明的点(i-1,j)...
			if (*src==keycolor)  
			{
				if(*(src-1)!=keycolor || 
				 *(src+1)!=keycolor || 
				 *(src+spitch)!=keycolor || 
				 *(src-spitch)!=keycolor)
					*dst=color;
			}
			else
			{
				*dst=*src;
			}

			dst++;
			src++;
		}
	}
}

extern _stdcall ablend_565_mmx(unsigned char alpha,unsigned char *lpSrc,
	long iSrcX, long iSrcY,long iSrcPitch, unsigned char *lpDst,long iDstX,
	long iDstY, long iDstPitch,long iDstW, long iDstH,WORD keycolor)
{
	//感谢::金点时空:: ---tiamo

	lpDst+=iDstY*iDstPitch+(iDstX<<1);
	lpSrc+=iSrcY*iSrcPitch+(iSrcX<<1);
	
	register __int64 ALPHA_MASK=MASKSHIFT*alpha;
	register __int64 ALPHA_KEYCOLOR=keycolor*MASKSHIFT;

	register long w=iDstW<<1;
	_asm
	{
		mov eax,iSrcPitch;
		sub eax,w
		mov ebx,iDstPitch;
		sub ebx,w

		mov esi,lpSrc;
		mov edi,lpDst;
		ALIGN 8;
		mov edx,iDstH;//H
START:
		mov ecx,iDstW;// iDstW/4
		shr ecx,2;
NEXT_POINT:
		
		movq mm0,[esi];		//Get Data
		movq mm1,[edi];

		movq mm5,ALPHA_MASK;
		movq mm6,mm1;
		pcmpeqw mm6,ALPHA_KEYCOLOR;
		pandn mm6,mm5;
		pxor mm7,mm7;

		movq mm2,mm0;//r
		pand mm2,MASKR;
		movq mm3,mm0;//g
		pand mm3,MASKG;
		movq mm4,mm0;//b
		pand mm4,MASKB;

		movq mm5,mm1;//b
		pand mm5,MASKB;
		psubw mm5,mm4;
		pmullw mm5,mm6;
		psrlw mm5,8;
		paddsw mm4,mm5;
		pand mm4,MASKB;
		paddusw mm7,mm4;

		movq mm5,mm1;//g
		pand mm5,MASKG;
		psrlw mm5,5;
		psrlw mm3,5;
		psubw mm5,mm3;
		pmullw mm5,mm6;
		psrlw mm5,8;
		paddsw mm3,mm5;
		psllw mm3,5;
		pand mm3,MASKG;
		paddusw mm7,mm3;

		movq mm5,mm1;//r
		pand mm5,MASKR;
		psrlw mm5,11;
		psrlw mm2,11;
		psubw mm5,mm2;
		pmullw mm5,mm6;
		psrlw mm5,8;
		paddsw mm2,mm5;
		psllw mm2,11;
		pand mm2,MASKR;
		paddusw mm7,mm2;

		movq [esi],mm7;
		add esi,8;
		add edi,8;
		dec ecx;
		jnz NEXT_POINT;
//NEXT_LINE:
		add esi,eax;
		add edi,ebx;

		dec edx;
		jnz START;
//DONE:
		EMMS;
	}
}

extern _stdcall ripplespread(long *lpbuf, long *lpoldbuf,long w,long h)
{
	//波的传播
	int i,j;
	long dataoffset=0;
	//为了防止越界1->h-1 1->w-1
	for(j=1;j<h-1;j++)
	{
		dataoffset=j*w;
		for(i=1;i<w-1;i++)
		{
			*(lpbuf+dataoffset)=(*(lpoldbuf+dataoffset-1)+*(lpoldbuf+dataoffset+1)
				+*(lpoldbuf+dataoffset-w)+*(lpoldbuf+dataoffset+w))/2-*(lpbuf+dataoffset);
			*(lpbuf+dataoffset)-=*(lpbuf+dataoffset)>>6;
			dataoffset++;
		}
	}
}
extern _stdcall renderipple(WORD *lpscreen,long screenpitch,WORD *lpbmp,long bmppitch,long *lpbuf,long w,long h)
{
	//
	int i,j,offsetx=0,offsety=0;
	long offsetdata,pos1,pos2;
	for(j=1;j<h-1;j++)
	{
		offsetdata=j*w;
		for(i=1;i<w-1;i++)
		{
			//
			offsetx=*(lpbuf+offsetdata+1)-*(lpbuf+offsetdata-1);
			offsety=*(lpbuf+offsetdata+w)-*(lpbuf+offsetdata-w);
			offsetx=offsetx>>2;
			offsety=offsety>>2;
			offsetdata++;
			if ((i+offsetx)>(w-1)) continue;
			if ((i+offsetx)<0) continue;
			if ((j+offsety)>(h-1)) continue;
			if ((j+offsety)<0) continue;
			pos1=j*screenpitch+i;
			pos2=(j+offsety)*bmppitch+(i+offsetx);
			*(lpscreen+pos1)=*(lpbmp+pos2);
		}
	}

}
extern _stdcall blur_c(WORD *lpscreen,long screenpitch,long x,long y,long w,long h)
{
	//模糊
	int i,j;
	WORD srccolor,r0,g0,b0;
	screenpitch=screenpitch>>1;
	lpscreen+=y*screenpitch+x;
	for(j=0;j<h;j++)
	{
		for(i=0;i<w;i++)
		{
			srccolor=*(lpscreen-1);
			r0=srccolor>>11;
			g0=(srccolor>>5)&0x3f;
			b0=srccolor&0x1f;

			srccolor=*(lpscreen+1);
			r0+=srccolor>>11;
			g0+=(srccolor>>5)&0x3f;
			b0+=srccolor&0x1f;
			
			srccolor=*(lpscreen-screenpitch);
			r0+=srccolor>>11;
			g0+=(srccolor>>5)&0x3f;
			b0+=srccolor&0x1f;

			srccolor=*(lpscreen+screenpitch);
			r0+=srccolor>>11;
			g0+=(srccolor>>5)&0x3f;
			b0+=srccolor&0x1f;
			
			r0=r0>>2;
			g0=g0>>2;
			b0=b0>>2;

			if(r0>0x1f) r0=0x1f;
			if(g0>0x3f) g0=0x3f;
			if(b0>0x1f) b0=0x1f;
			
			*lpscreen=(r0<<11)+(g0<<5)+b0;
			lpscreen++;
		}
		lpscreen+=screenpitch-w;
	}

}


extern _stdcall blur_mmx(char *lpscreen,long screenpitch,char *lpbmp,long bmppitch,long w,long h)
{
	//模糊效果*lpscreen(x,y)=(*lpbmp(x-1,y)+*lpbmp(x+1,y)+*lpbmp(x,y-1)*lpbmp(x-1,y+1))>>2
	//注意实际代码*lpscreen(x,y)=(2**lpbmp(x,y)+*lpbmp(x,y-1)+*lpbmp(x,y+1))/4
	register __int64 BLUR_MASK1=0xf7def7def7def7de;
	register __int64 BLUR_MASK2=0xe79ce79ce79ce79c;
	_asm
	{
		mov ebx,w
		add ebx,ebx
		mov esi,lpbmp
		add esi,8
		add esi,bmppitch

		mov edi,lpscreen
		add edi,8
		add edi,screenpitch

		mov edx,h
		sub edx,2
START:
		mov ecx,w			//字节计算
		shr ecx,2
		sub ecx,2			//2bytes

NEXT_POINT:
		mov eax,esi
		sub eax,ebx
		movq mm0,[esi]
		movq mm1,[esi+ebx]
		movq mm2,[eax]
		pand mm0,BLUR_MASK1
		pand mm1,BLUR_MASK2
		psrlw mm0,1
		pand mm2,BLUR_MASK2
		psrlw mm1,2
		psrlw mm2,2
		paddusw mm2,mm1
		paddusw mm0,mm2

		movq [edi],mm0
		add esi,8
		add edi,8
		dec ecx
		cmp ecx,0
		jnz NEXT_POINT
//NEXT_LINE:
		add esi,bmppitch
		sub esi,ebx
		add esi,16

		add edi,screenpitch
		sub edi,ebx
		add edi,16

		dec edx
		cmp edx,0
		jnz START
//DONE:
		EMMS
	}
}

extern _stdcall rotate_tran(WORD *lpscreen,long screenpitch,WORD *lpbmp,long bmppitch,long x,long y,long dw,long dh,long sx,long sy,long sw,long sh,float angle,WORD keycolor)
{
	//旋转:依次处理
	long i,j,tx,ty,x0,y0,centerx,centery;
	float sinangle=sin(angle),cosangle=cos(angle),cosangle0,sinangle0;
	long centersx,centersy;

	centerx=dw>>1;
	centery=dh>>1;

	WORD srccolor;
	bmppitch=bmppitch>>1;
	screenpitch=screenpitch>>1;

	lpscreen+=screenpitch*(y-centery)+x-centerx;
	
	centersx=sw>>1;
	centersy=sh>>1;
	//建立Sin Cos表
	for(i=0;i<dw;i++)
	{
		costablex[i]=(centerx-i)*cosangle;
		sintablex[i]=(centerx-i)*sinangle;
	}
	for(j=0;j<dh;j++)
	{
		costabley[j]=(centery-j)*cosangle;
		sintabley[j]=(centery-j)*sinangle;
	}

	for(j=0;j<dh;j++)
	{
		for(i=0;i<dw;i++)
		{
            x0=costablex[i]-sintabley[j]+centersx;
            y0=costabley[j]+sintablex[i]+centersy;
			if (x0>0 && x0<sw && y0>0 && y0<sh) 
			{
				srccolor=*(lpbmp+bmppitch*(y0+sy)+sx+x0);
				if(srccolor!=keycolor)
					*lpscreen=srccolor;
			}
			lpscreen++;
		}
		lpscreen+=screenpitch-dw;
	}
}

//以下代码参考
//extern _stdcall rotate_tran(WORD *lpscreen,long screenpitch,WORD *lpbmp,long bmppitch,long x,long y,long dw,long dh,long sx,long sy,long sw,long sh,float angle,WORD keycolor)
//{
	//旋转:依次处理
//	long i,j,tx,ty,x0,y0,centerx,centery;
//	float sinangle=sin(angle),cosangle=cos(angle),cosangle0,sinangle0;
//	long centersx,centersy;
//	float t0,t1,t2,t3;
//	WORD *lpscreen1,*lpscreen2,*lpscreen3;

//	centerx=dw>>1;
//	centery=dh>>1;

//	WORD srccolor;
//	bmppitch=bmppitch>>1;
//	screenpitch=screenpitch>>1;
//	lpscreen=lpscreen+screenpitch*(y-centery)+x-centerx;
	
//	centersx=sw>>1;
//	centersy=sh>>1;

//	lpscreen1=lpscreen+dw;
//	lpscreen2=lpscreen+dh*screenpitch;
//	lpscreen3=lpscreen2+dw;

//	for(j=0;j<=dh/2;j++)
//	{
//		for(i=0;i<=dw/2;i++)
//		{
			//
//			tx=centerx-i;
//			ty=centery-j;

//			t0=tx*cosangle;
//			t1=ty*sinangle;
///			t2=ty*cosangle;
//			t3=tx*sinangle;
//
//            x0=t0-t1+centersx;
//            y0=t2+t3+centersy;

//			if (x0>0 && x0<sw && y0>0 && y0<sh) 
//			{
//				srccolor=*(lpbmp+bmppitch*(y0+sy)+sx+x0);
//				if(srccolor!=keycolor)
//				{
//					*lpscreen=srccolor;
//				}
//			}
//
//            x0=-t0-t1+centersx;
//            y0=t2-t3+centersy;

//			if (x0>0 && x0<sw && y0>0 && y0<sh) 
//			{
//				srccolor=*(lpbmp+bmppitch*(y0+sy)+sx+x0);
//				if(srccolor!=keycolor)
//				{
//					*lpscreen1=srccolor;
//				}
//			}

//            x0=t0+t1+centersx;
//            y0=-t2+t3+centersy;

//			if (x0>0 && x0<sw && y0>0 && y0<sh) 
//			{
//				srccolor=*(lpbmp+bmppitch*(y0+sy)+sx+x0);
//				if(srccolor!=keycolor)
//				{
//					*lpscreen2=srccolor;
//				}
//			}

//            x0=-t0+t1+centersx;
//            y0=-t2-t3+centersy;

//			if (x0>0 && x0<sw && y0>0 && y0<sh) 
//			{
//				srccolor=*(lpbmp+bmppitch*(y0+sy)+sx+x0);
//				if(srccolor!=keycolor)
//				{
//					*lpscreen3=srccolor;
//				}
//			}


//			lpscreen++;
//			lpscreen1--;
//			lpscreen2++;
//			lpscreen3--;
//		}
//		lpscreen+=screenpitch-dw/2-1;
//		lpscreen1+=screenpitch+dw/2+1;
//		lpscreen2+=-screenpitch-dw/2-1;
//		lpscreen3+=-screenpitch+dw/2+1;
//	}
//}
extern _stdcall gray_565_mmx(char *lpscreen,long screenpitch,char *lpbmp,long bmppitch,long x,long y,long sx,long sy,long w,long h,WORD keycolor)
{
	//灰度转换（将彩色的图象转化为黑白图象r=g=b=(r+g+b)*3/8）
	long SrcOffset=sy*bmppitch+sx*2,
				  DstOffset=y*screenpitch+x*2;
	register __int64 ALPHA_KEYCOLOR=keycolor*MASKSHIFT;
	w=w*2;
	
	_asm
	{
		mov eax,bmppitch;
		mov ebx,screenpitch;
		mov esi,lpbmp;
		add esi,SrcOffset;
		mov edi,lpscreen;
		add edi,DstOffset;

		mov edx,h;   //H
		movq mm4,MASKB
START:
		mov ecx,w;// iDstW/4
		shr ecx,3;
NEXT_POINT:
		movq mm0,[esi]
		movq mm5,[edi]

		movq mm1,mm0
		movq mm2,mm0
		movq mm6,mm0

		psrlw mm1,6			//G
		psrlw mm2,11		//B
		
		pand mm0,mm4
		pand mm1,mm4
		pand mm2,mm4

		paddusw mm0,mm1
		paddusw mm0,mm2
		movq mm1,mm0

		psrlw mm0,2
		psrlw mm1,3
		paddusw mm0,mm1
		movq mm1,mm0
		pcmpgtw mm1,mm4
		por mm0,mm1
		pand mm0,mm4

		movq mm1,mm0
		psllw mm1,6
		
		movq mm2,mm0
		psllw mm2,11
		por mm0,mm1
		por mm0,mm2
		
		pcmpeqw mm6,ALPHA_KEYCOLOR
		pand mm5,mm6
		pandn mm6,mm0
		por mm5,mm6

		movq [edi],mm5
		add esi,8;
		add edi,8;
		dec ecx;
		cmp ecx,0;
		jnz NEXT_POINT;
//NEXT_LINE:
		add esi,eax;
		sub esi,w;
		add edi,ebx;
		sub edi,w;	

		dec edx;
		cmp edx,0;
		jnz START;
//DONE:
		EMMS;
	}
}

extern _stdcall blt_to_lighttable_mmx(unsigned char *lighttable,long iDstX,long iDstY,long w, long h,long iDstPitch,unsigned char *lpSrc,long iSrcX,long iSrcY,long iSrcPitch)
{
	//亮度信息取图像数据的 G通道(0-64)
	register __int64 FULL_LIGHT=MASKSHIFT*0xff;

	lpSrc+=iSrcPitch*iSrcY+(iSrcX<<1);
	lighttable+=iDstPitch*iDstY+iDstX;
	_asm
	{
		mov edi,lighttable
		mov esi,lpSrc
		movq mm6,FULL_LIGHT

		mov edx,h
START:
		mov ecx,w
		shr ecx,2

NEXT_POINT:
		movq mm0,[esi]
		pand mm0,MASKG
		psrlw mm0,3			//低8位为其亮度信息

		movd mm1,[edi]		//取出目标的亮度 4*8=32bits

		pxor mm7,mm7
		punpcklbw mm1,mm7   //扩展为64

		movq mm2,mm6
		movq mm3,mm6
		movq mm4,mm6

		psubusw mm2,mm0      //mm2=255-light1
		psubusw mm3,mm1      //mm3=255-light2
		
		pmullw mm2,mm3
		psrlw mm2,8
		
		psubusw mm4,mm2		//mm4即为结果
		

		movq mm7,mm4 //mm0
		psrlq mm7,24
		por mm7,mm4

		movd eax,mm7
		mov [edi],eax

		add edi,4
		add esi,8

		dec ecx
		jnz NEXT_POINT
		
		add edi,iDstPitch
		sub edi,w
		add esi,iSrcPitch
		sub esi,w
		sub esi,w
		
		dec edx
		jnz START

		emms
	}

}

extern _stdcall zbuffer_blt_mmx(unsigned char *lpscreen,unsigned char *lpzbuffer,long x,long y,WORD z,long scrw,unsigned char *lpsrc,long sx,long sy,long spitch,long sw,long sh,WORD keycolor)
{
	//Z_Buffer:WORD z;注意16bits
	sw&=0xfffc;

	register __int64 KEY_COLOR=MASKSHIFT*keycolor;
	register __int64 FULL_LIGHT=0xffffffffffffffff;
	register __int64 Z_BUFFER=MASKSHIFT*z;
	register __int32 offset=lpzbuffer-lpscreen;
	long tdw=(scrw-sw)<<1,tsw=spitch-(sw<<1);

	lpscreen+=y*(scrw<<1)+(x<<1);
	lpsrc+=sy*spitch+(sx<<1);
	lpzbuffer=lpscreen+offset;

	_asm
	{
		mov edi,lpscreen 
		mov eax,lpzbuffer
		mov esi,lpsrc

		movq mm7,KEY_COLOR
		movq mm6,Z_BUFFER
		mov ebx,sh     //loop h
		ALIGN 8

nextline:
		mov ecx,sw
		shr ecx,2
nextpoint:
		movq mm0,[esi]
		movq mm1,mm0
		pcmpeqw mm0,mm7

		//z_buffer begin
		movq mm2,[eax]		//mm1,src z_buffer
		movq mm3,mm2
		pcmpgtw mm2,Z_BUFFER
		movq mm5,mm2

		pxor mm2,mm0		//mm2=1：不透 mm0=1：透 mm2=(mm2 xor mm0) and mm2
		pand mm2,mm5
		movq mm4,mm2		//mm2=1:new_Z(需要作遮挡处理)  mm2=0:Old_Z
		
		pand mm2,mm6
		pandn mm4,mm3		//mm2=1使用Z_BUFFER替换原来的+透明 mm0
		por mm2,mm4
		movq [eax],mm2      //替换掉Z_Buffer

		pxor mm5,FULL_LIGHT	//mm0=(not mm5) or mm0
		por mm0,mm5
		//end
		
		psubusw mm1,mm0
		pand mm0,[edi]
		por mm0,mm1
		
		movq [edi],mm0

		add esi,8
		add edi,8
		add eax,8
		dec ecx
		jnz nextpoint

		add edi,tdw
		add eax,tdw
		add esi,tsw
		dec ebx
		jnz nextline
		emms
	}
}

extern _stdcall rle_blt(void *lpdst,long dpitch,long h,long x,long y,void *lpsrc,long pointernum)
{
	//rle blt
	//完成裁剪
	long sdx=x<<1,sdy=y,lenth,doffset=y*dpitch+(x<<1);
	_asm
	{
		mov edi,lpdst
		add edi,doffset
		mov esi,lpsrc
		mov ebx,pointernum
	    
nextPoint:
		mov	eax,[esi]
		add esi,4
		add edi,eax	//edi=edi+x
		add sdx,eax

		mov eax,[esi]
		add esi,4
		add sdy,eax
		mul dpitch
		add edi,eax		//edi=edi+y*dpith
		
		mov ecx,h
		cmp ecx,sdy
		jl endblt

		mov eax,[esi]
		mov lenth,eax
		add esi,4

		cmp sdy,0
		jl renext
		
		//memcpy frm esi to edi lenth*2 bytes
		mov ecx,lenth
		mov edx,sdx
		push edi
		
nextmem:
		//check eare
		add edx,2
		cmp edx,0
		jl addnext
		cmp edx,dpitch
		jnl addnext

		mov ax,[esi]
		mov [edi],ax
addnext:
		add edi,2
		add esi,2
		dec ecx
		jnz nextmem

		pop edi
		dec ebx
		jnz nextPoint
		jmp endblt
renext:
		add esi,lenth
		add esi,lenth
		dec ebx
		jnz nextPoint
endblt:
	}
}

extern _stdcall unsigned int RGB565(unsigned int RGB555)
{
	return (RGB555 & 0x1f)|((RGB555 & 0x7fe0)<<1);
}
