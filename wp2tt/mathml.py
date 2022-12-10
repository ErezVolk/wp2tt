"""Math conversion"""
import base64
import bz2
from io import BytesIO
from pathlib import Path

from lxml import etree


class Omml2Mathml:
    """Convert Office Math Markup Language to MathML"""
    MS_XSLT = Path("/Applications/Microsoft Word.app/Contents/Resources/omml2mathml.xsl")
    transform: etree.XSLT | None = None

    @classmethod
    def convert(cls, omml: etree._Entity) -> etree._ElementTree:
        """Convert Office Math Markup Language to MathML"""
        if cls.transform is None:
            cls.transform = cls._load_xslt()

        return cls.transform.apply(omml)

    @classmethod
    def _load_xslt(cls, try_ms=True) -> etree.XSLT:
        """Read transformer"""
        if try_ms and cls.MS_XSLT.is_file():
            dom = etree.parse(cls.MS_XSLT)
        else:
            encoded = TEIC_XSLT.encode()
            compressed = base64.a85decode(encoded)
            raw = bz2.decompress(compressed)
            dom = etree.parse(BytesIO(raw))

        return etree.XSLT(dom)

    @classmethod
    def make_buffer(cls, url: str) -> str:
        """Use this to convert the beta XSLT from TEIC."""
        import requests  # pylint: disable=import-outside-toplevel
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()

        raw = resp.content
        compressed = bz2.compress(raw, compresslevel=9)
        encoded = base64.a85encode(compressed, wrapcol=88)
        return encoded.decode()


if __name__ == "__main__":
    URL = "https://raw.githubusercontent.com/TEIC/Stylesheets/v7.54.0/docx/from/omml2mml.xsl"
    BUFFER = Omml2Mathml.make_buffer(URL)
    print(
        f'''# {URL} -> bz2 -> a85\n'''
        f'''TEIC_XSLT = r"""\n'''
        f'''{BUFFER}\n"""'''
    )


# https://raw.githubusercontent.com/TEIC/Stylesheets/v7.54.0/docx/from/omml2mml.xsl -> bz2 -> a85
TEIC_XSLT = r"""
6<\%_0gSqh;cuR`$!RBfLB%4%p\nuYs8W-!s8W,6s8W+Ks8W-!s8W+d"BYZ^1].4p!:iAj/7kM)OB$r#ei-6>:<d
GmRu3"M;JRk$2@Tt:<7Y_2.FXq!"1CoAJH4,c<<O3L"(jIB//e_G,21Z``XNJ]Ca(/4#>?K>!!!,Z!!"DJVn^bS1
J$B#lC1[0b+GP0dqKcKPu%?"<Al>"zzz!!>nh1'j.Er):iuc%M=MRYd"h1sf\@1`A+MTqKf16P.<=0`VC`1]V1k5
RfIWR5+shPgD]!9c>?G7&Q7!4\_^nVFH(R1hb8/!7N(;!75BZ)&aCnz!!!-N+sb!!-_o3f'_Ki[,>Et"@*p\/ab0
a0!!4to"DNZGJ,t3>@"9//^d\VU6r-`[e'a\1&./?5P%,_5ah8fdQi[K@8AYfS))d!WKEc*:!JiTTUl<bG12TJoa
<L]u6j*Y?U1"2C)P2L*r<>tmo`d,q9`WP]"(`N%1]RLU!$7it5QCca*5&3Up>o+#:os_n-k+7s9dM3E6::u"&cr[
W#gE^85Q`JZ8@8GK8;mt7X<O[A+tA73m6*oeK:rE],aK;lM2EXSBUU?BUm`VCJLF\?&Ys+R+XogU:`t='"<`A["*
mW&FHh`cbYWp>Pn4"IB`Y(/SA[[g:nVN',t+S56n(A;DE)>R5UdTUJ=lstr!ir,[[dY_rJsp)D23^`-j;+6'N@@u
5bu:`82S]3nI3s&g,#;?R.d+%RX.,0[J&CH#^I9siXl0P@&1'>T[LQglf!\b"HZ-4C+Oo-FBZd1X<AeIo91Wrq.\
<7O(E*@50Q$ZK;AS@7s4qZMQD=6WLlrA%Z:..+(P"lH1Tm7BplPW\)B$W"[I$_VuUub,M8VPPYn4e"9`@r<DKo#Y
2rVAp7NSi<#_E.6m957,1*2Fp&Y"YbLET4c"\sWNs#]H=+UKA.PX,H%K>8,kdTI:!^fjfKP?>GJk6%'p3-5=!(aN
8g*W]85oo`g4AtMK*]bi;=UOOUo=>d2@5T%"EMqGUoR>qMdn2TC*o#+3pli0_?G1YLL&/'d!e!:K+/s>5i:eu\#m
(:-FnT7Ho:[r0R@Z8(JN//jIpDLX)u;Jhjt)AXib!sQIeA0mb]p!_a"m_q,M4ftn)3IkBQTk.L[R2?m.A?[Pd.$'
ng@.Hj<;Hu)B_T20iK$$`9nLJ"R>XUoJt))?nZK.)G%B=X3Z0m=R!A":=jl1RB$hD$1M956csS4i#sA4%ar&c]gj
,rm'5$[OUelA4W1'bC`DJ><Jq@Qh+$??Z`X(=gN@@G'CIrM[fS2oLV[9YA=JmrcFM=g0)K'C8<'7DCZl:*rXikmF
SL(=26_+t(l^Lt_7r@l'Bd2+Z;6cV:o6)s=O9UfT0;nV-ICIFk09Jf4'mt4/nm2:Uu]7S@hR/8q=%$H*R66M8E6U
CfW)K;3d=USWoiE?1E.3X84#[7*;:f9L7MiUlK-25qsQLaj*pEupi/5Pn[U6[KgW:^8n%uk8m4QL8-4@9$5OtWA4
Ynja<Cb;RS]@W('[;:'p'%2'V,o"1-2`n_bfW7G!gLlK;VF<inQ?#&?qlT&jtL1[[kSIl,OgSl36qBhqd>CRE'M:
Nqe$=3)[cpHB6MI:%^>)H0^eNFjCs&HF\+aPTtpS#Wh%0JLuZEKSNbNSGKV\OmH2^(Npb7WAaJg(JLLMO_>po`%]
SsO@,kn+DK",1qE[o_@W+j8-B=494O"ZacZP*5kT:rHcVX,7NUuWObqCgE6DHlX968"Lld.m/0I@)1(u@%'dcIh9
)2__q"Zeu[P1H5Z)q']ST`7t_8?aUad_NHD]DJGD&KUu:7e3!W5[I"C$*fVP>;'8UV^G.a]""uH>2KTgh%>)QO.;
Km>k.g`&kd#aAKri1"S$Z=)=E2c=<D/;8.=DL==H7:3Q<nY=B+D:0]WVe&1l$,lDJg"OCsc?O#?&LLJ-ZBtu\6@W
,>"OTU1q#)54iL58*dBspYk`/JoE9SBk5@L_+TX>t'I%DCiEV?8%l@ceLS3)JZ(![r;A^0O6(;dWm>*&eloX.R;9
?Pb]N[<+9Z`7]4^(q&_"[dS[dZodl:=6cWLC[%\l_VLZ7"Q)LL_:tkH/ToBFPBN9e/-bL8PHHNU=9Y-OOEm2UYRY
'[jtAGGf50%=%N=!M'Wr'F3%^POK(^73YDpr=A"A$C//]bg1c@-?4-JbnhkP(jF%8s.e=`D.JES^-Tnc.RY-(-fN
7I*+lI&!3m/d'Y#um#1\t0]jcS^Y^eI8@*4!q4mlY8i9Vum8CdWO?s`TjkAEAenepYUi1(c9!abp_7`N,2TF`Pjk
[#7,MJJSj(TgI`$P-K1'1G&?5E3cWeDa(EAGBA;4S26"gFGpmZG<@^Z2Y&^u+H:q^e@[8NNWd9%CA37JpHc,J>50
a#6HMi(^p8!%[fLuPjM;HDeV@;tKGol3L7.<>Q9B,jr]E#Ua;3@_)$4JJ@$Qq-T+>TW!JY<DdK-lb>L)Mq7L)#eQ
";FDX7Ka.)KoDBk?q*e#=ZG>+&OXfnQr0eU$k*OM!Q#_F8]"/TW2,"/V'==H$AKm;B9GhoX.c3*O++i&kOeSRL@0
]=k4Wuo"/C(+gCF1P5C_!#HLA2CD?2oGAu4JT*.A8fQ0l6k0I.\jlfO'+BZ&FD1kBQE!u`$3?+hqUk@JNIb/Sf,5
&T;4/4QK0D\(YE4$nT\5j:,9/"At_Ya[OZ2ATmRDu<j5MpBN_3*k=+3hKlPp'2t=";j'Sh_2mpl/1DfGPkfl5MOB
9.N#+g4)b)fe7dE6=2>&h'^=)E9qt4ac%>])"2Wu[8L0K=,aeAeN]ukI@\tcH"f\<(b0;mp;KL=tEYr5PB-GU+p9
!eNr-sU91Kt$OA)cb@S+E2g%l--i&87KP-@g[laJl*fOM^UWMtr.#r3'FGkZ.f8@)e>(OiU(%R.d,nH\F\D+OBJY
@alpG7fj_XKg+>p>(k-O37/-s'KF%9J<*OSeKQ?h[9[F"3^RP&Q;nZ!3o8_h:scUiBZ]!jB""+Dd]+m+Mak3sT<h
G0>0,E%"29Rh[DCqu3o%bajl$Bs\))-L66Q!oXN-.')Q"FKMm`fS[m\D'-(d,g>9%m*f7*,gSf*V%48]3I:MuUOI
\$AHDo/q,Vb7kh,`l<dbN0504$/(c9P2o]X;I>I8\NG+c*l5Lh<*ai9,!c"T5e484A=f"U!.O:n(lp)3mPE"ns[!
<hN*Z)(q3[>kmHGY&>9MV`upJ!I]2T5IIU,$i`%$BYPi?RNr_8(gjhG<\?rc>]E-f^/erGp(`%@h!T"sS_jtW1e%
8:48X3q<*`]JBT37t-o5(=%H0dU`k3L/'`Ro'WAd]ZOqEFLbp;h82iU+43UcM9-=8Nl03P\u<FTMt<X3V+XCEW9t
/9_\l6!4\'<hUecZ,"X#J6lF1E._fipj,.D"#`SY5fRLR`u!;r7`%?`n>_-7O62Q9HLNTiW3@bT+nY]pVO!9E'Z1
2U!hpaRLQa/uDA%hA:f0`Eo,#*>^X5.XD(HS=k!Nlck#eFR,GDF:V\h;t-9:QP'[V<o,Of=.*:c$@p%7VCfm40#C
Z"MP)N"CY&!FmPe(jNhX.62IoN[9mb"E'T9+o#ja]UsPK.sT_"N&-U+BZI"&4.':#Y5#%0Lli<"2)QN8-$`t,RRs
._B;CC5tD&WW"g(Em82@':M(s<rDUVc1L7^/ZOQZ-pLF0ri()Qe]LgFpHs$g,&kO@[,pd"=EJSg/Ecto@T$6mC0a
W^>n5)]2!lN0f,a2BS8F3269Y5_&]s6>=2j&I^l6>u?).B`t2M,`ir.*M5ds,Jm,;[G]7heTP!`]Z2C'oMG3?rDa
PC7@>T^3%uD+hti5_LD9d9J\RMB@8.)$/kNLQC?Bi#:ql-l4L<Z>;+3FZ(D_i`L_M[iV"0nX*B`I(b8[hh[a.\!S
0A4<<N=_spT`M934TLNmuc@:$mP+qjs5(n(Zu=J=ji8ZBk+`2,Q=]?R=1WJ7WE2$ed7f`Gh#,.mtEPY+VR#ac;;j
i3Mk%29<"(M\,=RX>[KDU?X607#d4M/[7dM_Ni-1p[AMm-'@2nQ'b]RR_Nt65^pj^trHG(W9$KjVSXl*))K9K&P<
XZQqr+&kG838^qQ1#rVks9h`J8nDm:S:bkl\W*%K$kl@8*__SJ0P'&"+4s29V$*$mHZ$S2V;.3Wohb&hQ$'IY'*L
Dd%6)aun_^n/A^*`SJ:f@Or`n<q1V2ZF/B6OPLq6D9(,_F)]KFKk3-rBU_-*GbLGI.i6=2#)B42>1)BaV//a@318
666b25WK8dmo&6N?+!?X=C/5c"5h@'"657O&#ek!8qJd"#d_6;'?BM;IMW'lac&PT(eCALUa:`A#p;hc849+*;?B
6U*i^-OcX4oSIVV2&%E8SL@RH2B#I:3;@_n!ZQe"qmSXPeU_]G#6,ed(-XI3l5@JY<JdE+\OfQ#UYRk8eb1O@tlY
CC0`Yp2HUfaJSCBotgbE`,mD8b/$"&=3BI.j?<lg_.IuaW3.uXCrk),V/'f(m.HZmX\'!jfZp7gQ#=_(0d[jHlPh
\-4TWXK>n>L6?mc]%fNl\K;ht7FG>p<[2FL0Ag!';U*q6Ik',E,acPr90jk-L-Yn;F-7!$?>%[75`>-?g;"=l7@0
dVe&0rA`.bGLT(a0L>$$K%U3)M@^BL4":>;Hb;\[9H(0.oTMP"a*YFlXmHS6]AcZDsq)\FAOU*.r,B0#7RVc=p%9
Xb,WO7h@D")BoN_RV<*YL5JkT.;28@Lc6Hq2%\%L#VZ_/;diE@k-W-JEX,@"KJXCC)6%0Ukb:-eiPf!>d+ZCl8YD
1jR7oNrCKNYa<E/9Xd@HQLVN/u_dAbN$CQ+M)jG))n0,13,)C)T5)!aYc,b^&EM'e]3J?l(#.SDQA:+;\@>W?]P$
CrB&8?SBnD'r&Y==\4:1-rA&4"tJf:5;/VSg)IO@FIc#OHTfE%)`DUHndjGf:4n]FniS/JAj2N!sX:Kail+diIjA
!g+%oA82LO(`2b^&/O$9l@`l&N'uhM2BD>?f#.*f^69c#7@#g$jGs`mh"^6QtF9mPW+US5Q5_@qrX+_pp-\<7n%j
C5o,aDt4R-Yl"lY%ADM'rujM.e-E5dZ^9:'ojbM9eCS$=gn7I1)hWq&iVbQ(B-gTFdc&dKNV/LdGP`[6U_&[`9$)
=d`FQG*/ML.%j0]$L8aX@.4^S,K:llJROfTBeeXi&2'+IG:<SH*j3jDp/)>cS<5]sO@^V#%LB7.B^+pr,\!Bb%g3
sO:O4lt)\V4#`AB$j$1]?!29LQlS2P,8H6/`8ctMe[\Ju$#@.0_,8%ssQ6-L!!bQA?.+)"NHe1rSd(a`]e-MSrF,
@;?8^!NM0QjMmpnI#OE_W.fX3>dr]GP=p]IIUEj'o3Jo,a*Y&#$-==JnN0#$R=H'mXq`$qm2YV1,8kT,<eL8H7^m
(Tmg^akQMZQ</=Nuf%BZdG<Aa%]AeVoOV$H'M>]/kB=51/H3$l#'RBo0BG^s;&fu:E6R3jCG;LA!!$98ko(8PFKl
4NnT\#,D!1F<e((K/D#It*1&:=SA)[HQh&EtQ<U/A3l9mBMbTh")VPX=!0'mlh]4q)obcSM3QdFiGJDmX]2HF)!a
3ViTfSj>L'jm!7bAlP-%?qsYkFSr<p\5K2;2?aT;J6kd?#`[*(o,lG1:r9A"&9UBK8JO<aq&gHs:BD*4l\Zt[X>N
er$uTR.qHgl^BFk#Hi/\L+-j7?B+qYq]ek#1R02P0fO@p#da?YkMnfWKIkI)(/*9FRC.Gb,Wq7c<V/c^QtMI5WlT
%,Be?jBC5RLgpJF-1fqc316fc@#BlNZXJXj`<>5"/-OChr.5W&d-<:6_ljh%9iW)=q;bX71VF/#'hP*49PZ2BKn2
L)"!Iu-\nEeUOm8&W)PQ1HcKh4?,cSTGR>h@-DmjM&=Oc")$p'o919HK4r=[V&Ou&]k@tQKaFk<RQ8?f,0T^lHL/
;DJ,9s7A,[,CSO^nWVYe2ra#Dh14K0L)[-ATVeeGf:Q'===^BQpIGkGGY*VI0]9MKTg6;cl%pKus]18oJP;;(0[T
@WC`3l2Dt6+AR]#dtm4/\RC.&WAXEI.pFfA"H2I]+#BROR>6%R.gI3<%EZtI"Eom26_=K=qRSC,+X3HrpTS.;C#q
Y)Z]#5:J/_lQ75uU\MH[IJ@jH_Q"UV-m()3tj>4gnfaF=g=F5Ud(cb^*8ncTV>9d6F(h;21P0(!mrXTJFREa[g69
`\8&Rn,St[1pZQ,,1O"S.`AQ%.)s%R&b+7O*FbmA7<L[a"!7)]o59p"MtnE_;=mTp`Aka0UW>Ojq,K?`^jZkH0>\
V-R^C4#AWZmbb>6Na&@?u!YD]0b@@LV=M>gkV]Be6hogT5S@WQ^+s0<]@LBDA)Sou"qemKja2%=Wk&fRA$K"%fXO
fD,g;edS%PATHiO@"E0kP:\Y).7bYVX@j!BJi+L*:;F,cR/A&l#,Y0\4n8#"im=QO7rB\h]5AKo"FRR)Q&1BoOFZ
mQ40!:_;XmB*W&))C8jdd6JroLa8?<2%\_])?U&Z^"Uf2BOE%VWIluQ2Mj`%-)>K/2hIJ&>c>m,QQ=dlm)^gg9X,
12!X^0K!Kp-1aGc9D&2"Po7(J;2.7FO#`;q_P8f<f%5SFcB!Z5&E777e[@?H\.)(PhXXX40AfFpPBKK.qQ&P7LM7
RAd666Wjn#8do&@9`nF82QMSa<o%q\Zu\FDLuc7(<^4i+hSJj4THHWL;_tR:kGKa*WHkh#'@Ab3:KcZYk6m[]6^`
9(HLVfH.gLX3TDrt9%*Ksqu2Ssf=\m!khF/C8uR*KLo^kINSH7OP>'nLEXjQCSXprM)4*/'_Am`p&glY?H&.C\$4
MLs1Iu(F;-MK8b2fQ:?Bs^s3#Ht`idCkj3Bs.q!b[g*"@2&c5debu--c`:Vs0#W+e'W""Hc7icYhet08("F1Tfq&
R[?XU<("JMQ!jbV#se`cO<Fo:,WO-oMCX6g*ltu:(X4#'TR5g1rP'=A&4f)1%>Nm5'e(Ccq)>l!d:LrY^u92]M`J
S>&_Gu<0Ji%d!<?C8nI*nAKBE`hYCWp/fhLQ)h*=65O!XMpcTdERf0:8F-LX&FaM0qWA+`[ITg1V%Df")1W)PR)-
ZUtg*G14XBnh8#*&M-<s,I@/3MQcqJcYurLktNdRP%4<+s>&/-"8dO0h@-&5_tPbLo?,cLkpkC3BL7.%^uP0Q:A)
E]o'ELeh^#C9-L5<$e`&(`ZFsEq!;Olj<"Lqjlnj'BEA5`#%^<5$#pk$=aJpq32i2jg]GP%9*!JGBR!,d!cn%+7s
?V#UEh#hU-hBe,cW"3eKgh;#;tNRr$;f6f^JQ>j+i6f&CX0-'=_`D2`,IW?-#lNe\In."V%PcA\%u/#BeB4Pg1HZ
_iaK\"j1X=!A$kt&MS6&2$ui[4gHH*mZO[S*sb,b3#r'Yl8@C+CO&M[is2U_^;BF/.g$l^&)R`N!_<D;1U,MKdj&
!R#4us4)o/a;i+\%"53-+,Q9_+;,u^fMK=M2j5XVNVmLl[LOGO'0J_t4;E+e!,"6iT>:PIseE5mY5Y?&8CPgH2Vn
t/\&(1q%X(`J+%cd!C4mbX&e@E]6U.JULteW'Ds'Hn'!3(k^<i-4u_htU#uiM];#+Tb+C!L_#+$Z[olC`W6m6eP7
>lXh+MVbhdaO[N<TYZ5o+bW"9848-1T47[7KIC(""dl[$W0c1U7+p&ThTjqY]BK3M+/41F]k^ZVu&/cijKlgEfL6
)D\jECb.9;I*n:Eg+QfB=D)K,BJ#7K<bT9r!N[Sr\f0JO:'."boK983oS^oasoZ_qX#el+fnQ`=31@B9:#ql7!aT
5rO(?@C2I!p_6SVQU!km&<Td"J4YmOe*Wgr%uXEU<Q^>=),ukfFt^__-7'"f]dCU6e"f:J1Wc9m#][+Z`o/+Z:@,
$KR30*1lWN!J!P(TA+AhnYJT8$.c:`H<2^8E60J,P+^2EbNRuZY!QU_L%Ne2f4csW`h@c<"l+05e]FiSgq$OFXr(
Smu<5>G\P_/Kan-*dm`JL$eAZ!eXCQ15(JJYHnYY3jl/e%5\f@mGj\Nk4U:&0qlu5a&s1>41k?@6>i1;Yg!VkFgj
_fo5.JGY:uP_!SaO)oTfi6t<D6SNSmo"t5N%9n_acWCA9A*)eIg4CV1k(dK4++_;i2Nb5UJ/1e$2_MBBYNJQdDQ9
n,>NdWIB&#Nb0Q\L'UW0_IlW#9:,.Z0Ujl=qB2QtC&cPLQq$5tm[iTh65SokQP(lu9?mHb51-BNfI%QX^SGiII.9
8C33F1A,fZ&7>[bR5odAn3u*;CrHfVjlrkG(=H>_i^4*c*Kh\!,FH453/2`>q5n=ai/#Gn_Z1A-d*K07(@2+/3m#
a$%HnbYO$uIF*GL9!rVtfc=?T):`3B;qj[W+'*BJGq'_7nY_W,[e!GF87C(:l>e?M)qmXa:%*?]lP_AX;gLhUO2_
>o6q+fK,^MAR1kMpWaC:O5uf;mHT(Y@lT::,du_!;p4*!U.;8#4h$QR2c`7Q)b!!pB^Z4(''F,,+_0NPOgk\W'fA
SiG^sjVg/m^r-&V*KI:T2Jt8f!9->;'OI3^f,HSi9\GYQ1pm<B4:T+h"00VcBP#DZ#c+iWW*#HEqA$]bKQ=n^TiJ
E1V(Q4PjCNFS!ZEsWLOe_DjG3>8*JF@bE\PUl8j+[2X%F3loInep"8p%d46'+Bo=<h4ti?Fd!"f@uE/92>hH4O^%
$X*<I2f$e4o`<k':7g:UJCnJo!^W8%hj(BW]!)c>Zs6Ukr-91`&3<2/ZdhLG71i4+<,5!OcEdsLPICu-UJ2oJKS3
p_'"^>1lukI8#aur.P(nL]Q),o2OeR^_MJ3-$JY[!nTIC2/LhQp""D!E[CfA`Tgk0Qf2%PNTMFME?W"l!:=Cd"[J
d6+:JiHdW#O&UaW&c'ibNIu7HL(jHp)@AVNqf1kH%j!5*70$:PiX$63k=V,.<O]?_[IW6L/a-PG-jMj):4$[6TdX
'S4-C2hJgCIdjLaq"L)8LBoHOA.Aj5rf:f31$,21W8QBKA:S>]a$To]VJA[D31J2k81Bu'$OPk0`;Wu#BE&@r0S-
aiN0T(LP@2oMK1i\m++LmaGM+BZWLRIcq@"jPq'I&+r>Se\<(e69L3hW0JWf_[p/IA]n7q\:#q@Q>B#XE_0"h!d`
2rTAeC$\O8jHc+i)QriX"h?$6J9d:!km?i=ieT6aR2mc(@HoO*D&BH-L(OI11m>k%#k&9AK7Ff1&%\>p/K>I'nBp
G0IC^`)R]/-4KF&1s%e("H6+)LP'Gs6C*[/`U;%#]W,3@nTe(frO=ME(gDbXqT6Y*X'<cZeFG9--"_^::2\ZTknn
"Hb/^!X\)1_MB5`#04jZcKr'g;ISO7X/W(43bCDCmKMrJbUD/6UDf.X[CRDl/AUs"*hea6'</gCcIr9__sV!.o?l
`@DHh"@Ec_90li_OWiE%U.#=n'MBZ6(MMa2+.Dl60(Bt#9C0[dfUs^TQ,Y;OLOOlA:*Rim/m>c.SW/@OHnn27j(R
Sk]K:F(j3GN91Yr(0Ok]RD@Y'2k`?4,6"O>2T*JV$o:Jd/#]A8"P,M&6;jO_%Zt!J!';!@o;tKOI6S*k!)A&4QE#
Ljt9WF\&.0#Xi/8Fn=m8d!J8Z)i_LhD8*SWDNe>LC`+5"ZO.1.<pTMm5Z8e0F6bd>4d8onLB@gdO*:,CW@C8>>6.
-b&2-@=2EZ)_N-9sf^bLDe"!X8884]tDM$Q]L\O4<a@L^lV8HK*I8Fc":('u6H%K+Fp7"J62rD7Fa54c)G(?AIj_
+.)]o=ErBVXSn=!O_>cL?AqRo9,=38>9Y_C@":[<Z8Dj`&3/$R4ii_f,@uXj5?ge9oZNtK8e__pa.".B<%qpU!fg
8.@ShURA/9"h*YY9','UJ^#kUAP$BRF+VuAU<'(@2:'U>YVXR\D-qoP3%BUG6'))kNO!_7u0LE</,3GZ7%V@u"TH
YLiBhCXd@ORscUkf#$-j^_6+XA6s+d$FEGa9(DMSPIPYuO(!.M>?,%Jt1ZmU#XuW29782.kDOir_re&>@X*LhrnM
@UcobfsKM`)"e\7[!$lb5am;emnqFl)He>cPQO`].#\F#9ZIJ?,oC@uJgM\9"=?]LQn![g:'>b]W2FZ7<D!s%dSS
C]'XU_+>o9TiT`cCV-]B0/_^'N_2f72cW_d/o5X=]`jA#hE&6A\>pZP#,.N3[TFg.RFWjrN=+j,Z*Ha<^o&'+A\/
dA.\'B[s[$;QNfHHV/F2rEDP]A&&/laLBW<siVOC/G+KMnFpg6V+LR7*24HbgFuBE97]JYTSr2M+3^G&FT`I7?@Y
6!/QJo&"$(8nDRDf4"OlDQ@M%U_hW=^kb21F#Z_RGH_J^:K+c;HCN1/Im.i@oqoE9i%"q`kBV?*^BQ=)]M2MQ";n
t/AZX6@'.`"Xf^MfRpQ8O)M%ak[*0=2a+$n*Qt-u+Q,o3k.4*ti6PL;G\#?3_J-a0As3WZ3..T9dp<a:U^bll2%5
J.kg'$0=Y>RtmTcjVtWE/e^6roM58Y8#OdfU,dN?QnuaIA^[/'=lC@-,`&?R<E')diK?bpK`J1VgX;2]HlknR('b
QdA[C-%[dqK%&8[hcDC$rH:lu4S:Y0YB24+0>Tr9Lb&0ON*Z8@`cOtm/m'!YCI5X7l]cn_glaMLe*JKb^W8q`!3U
(Rb\X=1:qZLhfA3RE$INrD.gJMK=5-)eWFd^"QCb8"b"H;J3s5c#O+nlQroVGmUB6P.W?a>uOB8$!W'XA439#nn'
Hs-3u91BRpRh0-!s]l)XfO6c_W^UrUce%4kiW3r@g8MA0>8L=7iXq"=_f&[oP<IYpeDgN7U[P4RjB,F/Ci.8g<,d
`lWgYq(fdO4pqnCqfnHM;T]Mcp`d0.l=dWjoDo[A(S?OoPJZ+sJUH!@6iigMOF)E4C^[NKmLD<8'nqHc"0G=V=:R
M72.N&J5(V:]LIs-)cWQP">WR"(AtFz0u9Dbl42C.8L((,9f.(-mrCr%NJG+5)f<%<D<g%E!!e$G,b8/&^]5uDRr
,B$,d^,gQERU@Rr,UgD77:4XHM"',a)@*S6@*)&kO[`gN%.ENt#oB][^"^<G+>;@[h5CgMOCO#pb[*,a)C/qtn/S
P"?q/YDcs.`\EDb!C+f(P"V,0q2FJo<E2nF%C)Gd3R2]7WiE)Xh:qu:?i\;$er)r?bb+q<@4a@E[bG&fD'Y:pX,k
3h1V91?3R(UKk;p'lcek.YgROF&&=d6iJ,fR)H#?Y,+sH-,Wi;eW8RD-G&J5/08L0HlG%0%2ci\-oaf;A@W-4!(,
"X&1-jb]cLma[!R@0J2(tH>E+sI!e%>9m>roPCK+u!XH_3lDU$q7:6@e[Rm0o%MPg+iX2lpQ@.2t;i:*/i\J^9-/
SVOCjd1sg$iY52k0RU5a%p>["05?:#L$f=bo1G^gC\Z]?aB,JMRL]AOtj!9dZ.4#,E,b9aT[3[m/GR+:Cbfm3Vbl
qSo_&se1-,['PP"?p^<8CCo$7Cpu8L1f$P"?p^8L6@Q'M/=e!56ATbfn;P".7iH`O&efD7>t6WiE)>$F35LbIX^R
WiFo$dNDC4P'FEu"VOB`P$SYTa5ha=rpRV"LYqj]p<DJ:]?-r]gSo5Tm8f5$]3JNhkO\Oeob0u)h7)^Zgq38mqOc
X/Xei.8>/]/q<E3&,ghiOmG&Hs83R2S]VOD#8H29J'?I-sD93`cCB6MMLlL<TRc.QpA:u$QOELWMBMI+ok]boYA(
r\IZkslNSSg%'q7PTmIX0/tO/7fMm72/Tf$*TU.;+217C=jgH;([Ht\t%WC=U6LWY3kI$^$g=)pUqk!<E3%!79*>
f<E2If>tE5!P`i&HWiDn$Fju@[(s+Lk@hU1_R3]d,PgYfF9ehT:(r[]M[B44Rk9FcF]'G8`k9F0_\^#8kZq!.%F.
F/f.0rHE!!+5Q,QIfF+UJGf8LOS$,a(t),"X&Q!b;9284$\a!!rp)!;F1q8\OS>hY4_J2ne[,hY4]1pV$$c+UB=U
#XT)U+U!2O#XT)mk)OMRITS:;kfG]0*P6!@Rln=VA^-4).s4st?+[,\]6<St$8aXQf'ALB[k;.CkNb!ln"t8OG6M
Ol:PO5D<E3*)QBmgNeC7L.X'4_p2EMnB@kclZKo(Qm(s1=+V`gb3Z^g29<E)su<E3$u<E3%!<J!h0,a(_a##\#<T
E7NE8O*F#,a*ZaP"?r"'-r`5!%(8_P_:@`!!&+6^]4?Iji`uBfl#WXG1+pA<ioQ5C=T@U*ic6?g1$Kh)Ar#)a,_8
<l&Q#W(bfI?Nug-TRVAXfBr7=sP<?FEEb*:U,%\qZ/$d>LZGali;Fr.,!b@Hc5Qo1SP(3k%8L4?L,a(b"-:oJI!)
/PI-ifY\P#)078L0r1A<KML"N_pPJ-hB1,leT(P"G^$8L0N#9TT5\#YY3uKG!g/+UCCGJi1(o@`=6@/5Jbr\[f"f
WiDk)/_daPhI^F"0'%n/Tcf;]'".
"""
