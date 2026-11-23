% csee-report.cls
% 华中科技大学研究生课程报告正文，《中国电机工程学报》格式。封面、封底不在本类里，用课程下发的 Word 文件。
% 编译：XeLaTeX（latexmk -xelatex main.tex）。字体用 Windows 自带的宋体、黑体、仿宋、楷体、
% Times New Roman、Arial。版式参数取自机报 Word 模板：
%   A4；上 3.2 cm，下 2.5 cm，左右 2.0 cm；栏间距 0.78 cm；页眉距页顶 2.0 cm；
%   每页 42 行（单倍行距对齐网格，行距约 16.2 pt）；正文五号宋体 / Times New Roman。
\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{csee-report}[2026/11/17 v1.2 HUST course report in Proceedings of the CSEE style]

% 1.2857 = 16.2 bp 行距 / ctex 五号的单倍行距 12.65 pt
\LoadClass[a4paper,twocolumn,zihao=5,fontset=windows,linespread=1.2857]{ctexart}

\RequirePackage[a4paper,top=3.2cm,bottom=2.5cm,left=2cm,right=2cm,
  headheight=26pt,headsep=21.5bp,footskip=0.5cm,columnsep=0.78cm]{geometry}
\RequirePackage{amsmath,amssymb}
\RequirePackage[math-style=ISO,bold-style=ISO]{unicode-math}
\RequirePackage{graphicx}
\RequirePackage{array,booktabs,multirow,tabularx}
\RequirePackage{caption}
\RequirePackage{enumitem}
\RequirePackage{fancyhdr}
\RequirePackage{siunitx}
\RequirePackage{needspace}
\RequirePackage{gbt7714}
\RequirePackage[hidelinks]{hyperref}

\graphicspath{{figures/}}

% ---------------------------------------------------------------- 字体
\setmainfont{Times New Roman}
\setsansfont{Arial}
\setmathfont{TeX Gyre Termes Math}
% 公式尺寸按模板（MathType 尺寸定义）：正文 10.5 pt，上下标 6.5 pt，次上下标 4.5 pt
\begingroup\normalsize\xdef\csee@normalfsize{\f@size}\endgroup
\DeclareMathSizes{\csee@normalfsize}{\csee@normalfsize}{6.52}{4.52}
% fontset=windows 的无衬线中文是微软雅黑；机报的题名、各级标题、“参考文献”都是黑体
\renewcommand\CJKsfdefault{zhhei}
% 汉字与西文之间不自动加空（模板：“式(3)”“晶闸管(integrated gate……)”括号紧贴汉字）；
% 源文件里写了空格的地方仍排一个 1/4 字宽的普通空格
\xeCJKsetup{CJKecglue={}}
% 固定行距：#1 字号，#2 行距（不受全局 linespread 影响）
\newcommand\csee@fixed[2]{\linespread{1}\fontsize{#1}{#2}\selectfont}
% 数字：自小数点起前后每 3 位留 1/4 字符空格，4 位数也分节（模板示例 0.099 7）；数字与单位间 1/4 字符
\sisetup{group-minimum-digits=4, group-separator={\hspace{0.25em}}, number-unit-separator={\hspace{0.25em}}}

% ---------------------------------------------------------------- 信息字段
\def\csee@title{}
\def\csee@author{}
\def\csee@affil{}
\def\csee@entitle{}
\def\csee@enauthor{}
\def\csee@enaffil{}
\def\csee@abstract{}
\def\csee@keywords{}
\def\csee@enabstract{}
\def\csee@enkeywords{}
\def\csee@fund{}
\def\csee@journalcn{中\quad 国\quad 电\quad 机\quad 工\quad 程\quad 学\quad 报}
\def\csee@journalen{Proceedings of the CSEE}
\def\csee@oddheader{}
\def\csee@articleno{}
\def\csee@clc{}
\def\csee@doccode{}
\def\csee@subject{}
\def\csee@bio{}

\renewcommand\title[1]{\def\csee@title{#1}}
\renewcommand\author[1]{\def\csee@author{#1}}
\newcommand\affil[1]{\def\csee@affil{#1}}
\newcommand\entitle[1]{\def\csee@entitle{#1}}
\newcommand\enauthor[1]{\def\csee@enauthor{#1}}
\newcommand\enaffil[1]{\def\csee@enaffil{#1}}
\renewcommand\abstract[1]{\def\csee@abstract{#1}}
\newcommand\keywords[1]{\def\csee@keywords{#1}}
\newcommand\enabstract[1]{\def\csee@enabstract{#1}}
\newcommand\enkeywords[1]{\def\csee@enkeywords{#1}}
\newcommand\fund[1]{\def\csee@fund{#1}}
\newcommand\journalheader[2]{\def\csee@journalcn{#1}\def\csee@journalen{#2}}
\newcommand\oddheader[1]{\def\csee@oddheader{#1}}
% 首页最上方一行：文章编号、中图分类号、文献标志码、学科分类号，各项可选，空的不排
\newcommand\articleno[1]{\def\csee@articleno{#1}}
\newcommand\clc[1]{\def\csee@clc{#1}}
\newcommand\doccode[1]{\def\csee@doccode{#1}}
\newcommand\subjectcode[1]{\def\csee@subject{#1}}
% 作者简介：\biography{照片文件}{姓名}{简介正文}，排在参考文献之后
\newcommand\biography[3]{\def\csee@bio{{#1}{#2}{#3}}}

% ---------------------------------------------------------------- 页眉
% 小五宋体；刊名居中，页码右对齐在 16.8 cm 制表位（离右边距 0.2 cm）或左对齐。
% 普通页：一行，下方 0.75 pt 单线。首页：两行（刊名 / 英文刊名），下方双细线，
% 位置按 Word 模板实测：普通页基线距页顶 65.8 bp，首页两行基线 65.8 / 77.1 bp。
\newcommand\csee@headfont{\csee@fixed{9bp}{11.3bp}}
\newcommand\csee@rpage{\llap{\thepage\hspace{0.2cm}}}
\newcommand\csee@firsthead{%
  \raisebox{-15.8bp}[0pt][0pt]{\parbox[b]{\headwidth}{\csee@headfont
    \hbox to \headwidth{\hfil\csee@journalcn\hfil}%
    \hbox to \headwidth{\hfil\csee@journalen\hfil\csee@rpage}%
    \vspace{0.65bp}\hrule height 0.5bp\vspace{0.5bp}\hrule height 0.5bp}}}
\fancypagestyle{csee@first}{%
  \fancyhf{}%
  \fancyhead[C]{\csee@firsthead}%
  \renewcommand\headrulewidth{0pt}}
\fancypagestyle{csee}{%
  \fancyhf{}%
  \fancyhead[L]{\csee@headfont\ifodd\value{page}\else\thepage\fi}%
  \fancyhead[C]{\csee@headfont\ifodd\value{page}\csee@oddheader\else\csee@journalcn\fi}%
  \fancyhead[R]{\csee@headfont\ifodd\value{page}\csee@rpage\fi}%
  \renewcommand\headrulewidth{0.75bp}}

% ---------------------------------------------------------------- 首页编号行
% 小五，项目名黑体、内容 Times New Roman，全角冒号，各项之间空 4 格（2 字宽）；
% 中图分类号的字母与数字之间写一个普通空格，Times New Roman 的空格正好 1/4 字宽
\def\csee@metasep{}
\newcommand\csee@metaitem[2]{%
  \ifx#2\@empty\else\csee@metasep{\heiti #1：}#2\def\csee@metasep{\hspace{2em}}\fi}
\newcommand\csee@metaline{%
  \ifx\csee@articleno\@empty\ifx\csee@clc\@empty\ifx\csee@doccode\@empty\ifx\csee@subject\@empty
    \let\csee@metaempty\relax\fi\fi\fi\fi
  \ifdefined\csee@metaempty\else
    {\noindent\csee@fixed{9bp}{16.2bp}%
      \csee@metaitem{文章编号}\csee@articleno
      \csee@metaitem{中图分类号}\csee@clc
      \csee@metaitem{文献标志码}\csee@doccode
      \csee@metaitem{学科分类号}\csee@subject\par}%
  \fi}

% ---------------------------------------------------------------- 题名与摘要
% 题名、作者、单位通栏排；中英文摘要和关键词排在左栏开头（与机报模板一致）
\renewcommand\maketitle{%
  \clearpage
  \pagenumbering{arabic}%
  \pagestyle{csee}%
  \twocolumn[{%
    \csee@metaline
    \begin{center}
      \vspace*{12bp}
      {\sffamily\csee@fixed{22bp}{30bp}\csee@title\par}
      \vspace{8bp}
      {\fangsong\csee@fixed{14bp}{22bp}\csee@author\par}
      \vspace{15bp}
      {\kaishu\csee@fixed{10.5bp}{16.2bp}（\csee@affil）\par}
      % 单位与英文题名之间空一行（Word 模板里这一行按网格占两行高）
      \vspace{52bp}
      {\bfseries\csee@fixed{12bp}{16.2bp}\csee@entitle\par}
      {\csee@fixed{10.5bp}{16.2bp}\csee@enauthor\par}
      {\csee@fixed{9bp}{14bp}(\csee@enaffil)\par}
    \end{center}
    \vspace{12bp}}]%
  \thispagestyle{csee@first}%
  % 基金项目：首页左栏底部的无编号脚注，六号字，首行缩进 2 字符。直接插入 \footins，
  % 不经过 \footnotetext，免得 hyperref 为空编号建锚点
  \ifx\csee@fund\@empty\else
    \insert\footins{%
      \interlinepenalty\interfootnotelinepenalty
      \splittopskip\footnotesep \splitmaxdepth\dp\strutbox \floatingpenalty\@MM
      \hsize\columnwidth \@parboxrestore
      \csee@fixed{7.5bp}{11bp}\parindent 2em
      \rule{\z@}{\footnotesep}\textbf{基金项目：}\csee@fund\par}%
  \fi
  \begingroup
    \csee@fixed{9bp}{14bp}\setlength\parindent{0pt}%
    \textbf{ABSTRACT:} \csee@enabstract\par
    \vspace{6bp}\textbf{KEY WORDS:} \csee@enkeywords\par
    \vspace{6bp}{\heiti 摘要：}\csee@abstract\par
    \vspace{6bp}{\heiti 关键词：}\csee@keywords\par
  \endgroup
  \vspace{6bp}}

\renewcommand\footnoterule{\kern-3pt\hrule width 0.4\columnwidth\kern 2.6pt}

% ---------------------------------------------------------------- 标题
% 一级：小四黑体 / Arial，段前段后 6 pt；二级：五号黑体 / Arial；三级：五号宋体 / Times New Roman
% 换行不悬挂缩进
\ctexset{
  section = {
    format = \sffamily\csee@fixed{12bp}{16.2bp}\raggedright,
    aftername = \hspace{1em},
    beforeskip = 6bp, afterskip = 6bp, hang = false,
  },
  subsection = {
    format = \sffamily\csee@fixed{10.5bp}{16.2bp}\raggedright,
    aftername = \hspace{0.5em},
    beforeskip = 0pt, afterskip = 1sp, hang = false,
  },
  subsubsection = {
    format = \rmfamily\csee@fixed{10.5bp}{16.2bp}\raggedright,
    aftername = \hspace{1em},
    beforeskip = 0pt, afterskip = 1sp, hang = false,
  },
}

% 附录：一级标题为“附录 A”，图表公式编号为 A1、A2……
\renewcommand\appendix{%
  \par
  \setcounter{section}{0}%
  \renewcommand\thesection{\Alph{section}}%
  \ctexset{section/name = {附录\enspace,}, section/aftername = {}}%
  \counterwithin*{figure}{section}%
  \counterwithin*{table}{section}%
  \counterwithin*{equation}{section}%
  \renewcommand\thefigure{\thesection\arabic{figure}}%
  \renewcommand\thetable{\thesection\arabic{table}}%
  \renewcommand\theequation{\thesection\arabic{equation}}}

% 致谢：标题居中
% 标题后禁止断页，免得“致谢”落在栏底而内容在下一栏
\newenvironment{acknowledgment}{%
  \par\addvspace{6bp}
  {\centering\heiti\csee@fixed{10.5bp}{16.2bp}致\quad 谢\par}
  \nobreak\vspace{6bp}\@afterheading}{\par}

% ---------------------------------------------------------------- 列表
% 1）工作模式 1。编号缩进 2 字符，换行顶格
\setlist{nosep}
\setlist[enumerate]{label=\arabic*）, wide=2em, labelsep=0pt}
\setlist[itemize]{wide=2em}

% ---------------------------------------------------------------- 图表
% 图题、表题：中文小五黑体，数字和英文 Times New Roman 加粗，行距 14 pt
% \bicaption{中文题名}{English caption}：先排中文题名，下一行排英文题名
\captionsetup{font={bf}, labelsep=quad, justification=centering,
  singlelinecheck=false, skip=3bp}
\DeclareCaptionFont{csee}{\csee@fixed{9bp}{14bp}}
\captionsetup{font+=csee}
\captionsetup[figure]{position=below, aboveskip=3bp, belowskip=0pt}
\captionsetup[table]{position=above, aboveskip=0pt, belowskip=0pt}
\newcommand\csee@en@figure{Fig.}
\newcommand\csee@en@table{Tab.}
\newcommand\bicaption[2]{%
  \caption{#1}%
  {\centering\bfseries\csee@fixed{9bp}{14bp}%
    \csname csee@en@\@captype\endcsname\ \csname the\@captype\endcsname\quad #2\par}%
  \ifx\@captype\csee@str@table\vspace{3bp}\fi}
\def\csee@str@table{table}

% 表格：三线表，顶线底线 0.75 pt，栏目线 0.5 pt；表内六号字，行距 14 pt
\setlength\heavyrulewidth{0.75pt}
\setlength\lightrulewidth{0.5pt}
\setlength\cmidrulewidth{0.5pt}
% 浮动体内部会被 \@floatboxreset 重置为 \normalsize，所以字号要加在它后面
\g@addto@macro\@floatboxreset{\ifx\@captype\csee@str@table\csee@fixed{7.5bp}{14bp}\fi}
% 三线表铺满栏宽（在 table* 里为版心宽），多出的宽度均分到各列间，与机报示例表一致：
%   \begin{fulltabular}{ccccc} ... \end{fulltabular}
\newenvironment{fulltabular}[1]
  {\begin{tabular*}{\linewidth}{@{\extracolsep{\fill}\hspace{\tabcolsep}}#1}}
  {\end{tabular*}}
\setlength\textfloatsep{8bp plus 2bp minus 2bp}
\setlength\floatsep{8bp plus 2bp minus 2bp}
\setlength\intextsep{8bp plus 2bp minus 2bp}

% ---------------------------------------------------------------- 公式
\setlength\abovedisplayskip{4bp plus 2bp minus 2bp}
\setlength\belowdisplayskip{4bp plus 2bp minus 2bp}

% ---------------------------------------------------------------- 参考文献
% 顺序编码制，正文上标引用；小五宋体，行距 14 pt。著录标点按机报模板用全角“，：．”（括号半角），
% 中文文献可用 translation 字段在下一行接英文译文，见 csee-gbt7714.bst
\bibliographystyle{csee-gbt7714}
\@namedef{bibstyle@csee-gbt7714}{\bibstyle@super}
\setlength\bibsep{0pt}
% 著录项之间是全角句点“．”，其后不再加空（natbib 的 \newblock 默认加一段间距）
\renewcommand\bibfont{\csee@fixed{9bp}{14bp}\def\newblock{\unskip\ignorespaces}}
% 悬挂缩进固定 0.63 cm，编号左对齐，与编号宽度无关（natbib 默认按最宽编号定缩进）
\renewcommand\NAT@bibsetnum[1]{%
  \setlength\leftmargin{0.63cm}\setlength\labelwidth{0.63cm}%
  \setlength\labelsep{0pt}\setlength\itemindent{0pt}%
  \setlength\itemsep{\bibsep}\setlength\parsep{\z@}%
  \def\makelabel##1{##1\hfil}}
% “参考文献”：小四黑体，段前段后 4 pt，左对齐（按模板实测）
\renewcommand\bibsection{%
  \par\addvspace{4bp}{\noindent\sffamily\csee@fixed{12bp}{16.2bp}\refname\par}\nobreak\vspace{4bp}}

% ---------------------------------------------------------------- 作者简介
% 参考文献之后：4 cm、0.75 pt 短线；左侧照片（宽 2.2 cm、高 2.8 cm）下排姓名，右侧“作者简介：”
% 和简介正文；小五宋体 / Times New Roman，行距 14 pt，首行缩进 0.63 cm
\def\csee@makebio@#1#2#3{%
  \par\vspace{6bp}\Needspace{3.6cm}%
  \noindent\rule{4cm}{0.75bp}\par\vspace{3bp}%
  \begingroup\csee@fixed{9bp}{14bp}%
  \noindent\begin{minipage}[t]{2.2cm}\vspace{0pt}\centering
    \includegraphics[width=2.2cm,height=2.8cm]{#1}\par\vspace{1bp}#2\par
  \end{minipage}\hspace{0.1cm}%
  \begin{minipage}[t]{\dimexpr\linewidth-2.3cm\relax}\vspace{0pt}%
    \setlength\parindent{0.63cm}\indent{\heiti 作者简介：}\par\indent #3\par
  \end{minipage}\par
  \endgroup}
\AtEndDocument{\ifx\csee@bio\@empty\else\expandafter\csee@makebio@\csee@bio\fi}

\endinput
