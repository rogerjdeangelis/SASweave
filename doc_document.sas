%let pgm=doc_document;

%let purpose=Documenting your programs with a poor mans knitr or SASweave;

* Free BoxOft [df to ppt ( see vba code to add your theme to ppt slides;
%let pdf2ppt=d:\exe\p2p\pdftopptcmd.exe;  * free boxoft pdf to ppt converter;

* OUTPUTS;
%let wevoutpdf=d:\ymr\&pgm..pdf;          * output pdf;
%let wevoutppt=d:\ymr\&pgm..ppt;          * output pdf;

******************************************************************************;

proc datasets library=work kill;
run;quit;;

%Macro utl_ymrlan100
    (
      style=utl_ymrlan100
      ,frame=box
      ,TitleFont=14pt
      ,docfont=13pt
      ,fixedfont=12pt
      ,rules=ALL
      ,bottommargin=.25in
      ,topmargin=.25in
      ,rightmargin=.25in
      ,leftmargin=.25in
      ,cellheight=13pt
      ,cellpadding = 2pt
      ,cellspacing = .2pt
      ,borderwidth = .2pt
    ) /  Des="SAS PDF Template for PDF";

ods path work.templat(update) sasuser.templat(update) sashelp.tmplmst(read);

proc template ;
source styles.printer;
run;quit;

Proc Template;

   define style &Style;
   parent=styles.rtf;

        class body from Document /

               protectspecialchars=off
               asis=on
               bottommargin=&bottommargin
               topmargin   =&topmargin
               rightmargin =&rightmargin
               leftmargin  =&leftmargin
               ;

        class color_list /
              'link' = blue
               'bgH'  = _undef_
               'fg'  = black
               'bg'   = _undef_;

        class fonts /
               'TitleFont2'           = ("Arial, Helvetica, Helv",&titlefont,Bold)
               'TitleFont'            = ("Arial, Helvetica, Helv",&titlefont,Bold)

               'HeadingFont'          = ("Arial, Helvetica, Helv",&titlefont)
               'HeadingEmphasisFont'  = ("Arial, Helvetica, Helv",&titlefont,Italic)

               'StrongFont'           = ("Arial, Helvetica, Helv",&titlefont,Bold)
               'EmphasisFont'         = ("Arial, Helvetica, Helv",&titlefont,Italic)

               'FixedFont'            = ("Courier New, Courier",&fixedfont)
               'FixedEmphasisFont'    = ("Courier New, Courier",&fixedfont,Italic)
               'FixedStrongFont'      = ("Courier New, Courier",&fixedfont,Bold)
               'FixedHeadingFont'     = ("Courier New, Courier",&fixedfont,Bold)
               'BatchFixedFont'       = ("Courier New, Courier",&fixedfont)

               'docFont'              = ("Arial, Helvetica, Helv",&docfont)

               'FootFont'             = ("Arial, Helvetica, Helv", 9pt)
               'StrongFootFont'       = ("Arial, Helvetica, Helv",8pt,Bold)
               'EmphasisFootFont'     = ("Arial, Helvetica, Helv",8pt,Italic)
               'FixedFootFont'        = ("Courier New, Courier",8pt)
               'FixedEmphasisFootFont'= ("Courier New, Courier",8pt,Italic)
               'FixedStrongFootFont'  = ("Courier New, Courier",7pt,Bold);

        class GraphFonts /
               'GraphDataFont'        = ("Arial, Helvetica, Helv",&fixedfont)
               'GraphValueFont'       = ("Arial, Helvetica, Helv",&fixedfont)
               'GraphLabelFont'       = ("Arial, Helvetica, Helv",&fixedfont,Bold)
               'GraphFootnoteFont'    = ("Arial, Helvetica, Helv",8pt)
               'GraphTitleFont'       = ("Arial, Helvetica, Helv",&titlefont,Bold)
               'GraphAnnoFont'        = ("Arial, Helvetica, Helv",&fixedfont)
               'GraphUnicodeFont'     = ("Arial, Helvetica, Helv",&fixedfont)
               'GraphLabel2Font'      = ("Arial, Helvetica, Helv",&fixedfont)
               'GraphTitle1Font'      = ("Arial, Helvetica, Helv",&fixedfont)
               'NodeDetailFont'       = ("Arial, Helvetica, Helv",&fixedfont)
               'NodeInputLabelFont'   = ("Arial, Helvetica, Helv",&fixedfont)
               'NodeLabelFont'        = ("Arial, Helvetica, Helv",&fixedfont)
               'NodeTitleFont'        = ("Arial, Helvetica, Helv",&fixedfont);


        style Graph from Output/
                outputwidth = 100% ;

        style table from table /
                outputwidth=100%
                protectspecialchars=off
                asis=on
                background = colors('tablebg')
                frame=&frame
                rules=&rules
                cellheight  = &cellheight
                cellpadding = &cellpadding
                cellspacing = &cellspacing
                bordercolor = colors('tableborder')
                borderwidth = &borderwidth;

         class Footer from HeadersAndFooters

                / font = fonts('FootFont')  just=left asis=on protectspecialchars=off ;

                class FooterFixed from Footer
                / font = fonts('FixedFootFont')  just=left asis=on protectspecialchars=off;

                class FooterEmpty from Footer
                / font = fonts('FootFont')  just=left asis=on protectspecialchars=off;

                class FooterEmphasis from Footer
                / font = fonts('EmphasisFootFont')  just=left asis=on protectspecialchars=off;

                class FooterEmphasisFixed from FooterEmphasis
                / font = fonts('FixedEmphasisFootFont')  just=left asis=on protectspecialchars=off;

                class FooterStrong from Footer
                / font = fonts('StrongFootFont')  just=left asis=on protectspecialchars=off;

                class FooterStrongFixed from FooterStrong
                / font = fonts('FixedStrongFootFont')  just=left asis=on protectspecialchars=off;

                class RowFooter from Footer
                / font = fonts('FootFont')  asis=on protectspecialchars=off just=left;

                class RowFooterFixed from RowFooter
                / font = fonts('FixedFootFont')  just=left asis=on protectspecialchars=off;

                class RowFooterEmpty from RowFooter
                / font = fonts('FootFont')  just=left asis=on protectspecialchars=off;

                class RowFooterEmphasis from RowFooter
                / font = fonts('EmphasisFootFont')  just=left asis=on protectspecialchars=off;

                class RowFooterEmphasisFixed from RowFooterEmphasis
                / font = fonts('FixedEmphasisFootFont')  just=left asis=on protectspecialchars=off;

                class RowFooterStrong from RowFooter
                / font = fonts('StrongFootFont')  just=left asis=on protectspecialchars=off;

                class RowFooterStrongFixed from RowFooterStrong
                / font = fonts('FixedStrongFootFont')  just=left asis=on protectspecialchars=off;

                class SystemFooter from TitlesAndFooters / asis=on
                        protectspecialchars=off just=left;
    end;
run;
quit;

%Mend utl_ymrlan100;
%utl_ymrlan100;

%Macro Tut_Sly
(
 stop=43,
 L1=' ',    L2=' ', L3=' ', L4=' ', L5=' ', L6=' ', L7=' ', L8=' ', L9=' ',
 L10=' ', L11=' ',
 L12=' ', L13=' ', L14=' ', L15=' ', L16=' ', L17=' ', L18=' ', L19=' ',
 L20=' ', L21=' ',
 L22=' ', L23=' ', L24=' ', L25=' ', L26=' ', L27=' ', L28=' ', L29=' ', L30=' ', L31=' ', L32=' ',
 L33=' ', L34=' ', L35=' ', L36=' ', L37=' ', L38=' ', L39=' ', L40=' ', L41=' ', L42=' ',L43=' ',
 L44=' ', L45=' ', L46=' ', L47=' ', L48=' ', L49=' ', L50=' ', L51=' ', L52=' '
 )/ des="SAS Slides all argument values need to be single quoted";

/* creating slides for a presentation */
/* up to 32 lines */
/* backtic ` is converted to a single quote  */
/* | is converted to a , */

Data _OneLyn1st(rename=t=title);

Length t $255;
 t=resolve(translate(&L1,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
 t=resolve(translate(&L2,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
 t=resolve(translate(&L3,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
 t=resolve(translate(&L4,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
 t=resolve(translate(&L5,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
 t=resolve(translate(&L6,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
 t=resolve(translate(&L7,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
 t=resolve(translate(&L8,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
 t=resolve(translate(&L9,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L10,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L11,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L12,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L13,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L14,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L15,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L16,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L17,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L18,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L19,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L20,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L21,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L22,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L23,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L24,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L25,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L26,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L27,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L28,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L29,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L30,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L31,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L32,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L33,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L34,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L35,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L36,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L37,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L38,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L39,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L41,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L42,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L43,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L44,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L45,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L46,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L47,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L48,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L50,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L51,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;
t=resolve(translate(&L52,"'","`"));t=translate(t,",","|");t=translate(t,";","~");t=translate(t,'%',"#");t=translate(t,'&',"@");Output;

run;quit;

/*  %let l7='^S={font_size=25pt just=c cellwidth=100pct}Premium Dollars';  */

options label;
%if &stop=7 %then %do;
   data _null_;
      tyt=scan(&l7,2,'}');
      call symputx("tyt",tyt);
   run;
   ods pdf bookmarkgen=on bookmarklist=show;
   ods proclabel="&tyt";run;quit;
%end;
%else %do;
   ods proclabel="Title";run;quit;
%end;


data _onelyn;
  set _onelyn1st(obs=%eval(&stop + 1));
  if not (left(title) =:  '^') then do;
     pre=upcase(scan(left(title),1,' '));
     idx=index(left(title),' ');
     title=substr(title,idx+1);
  end;
  put title;
run;

* display the slide ;
title;
footnote;

proc report data=_OneLyn nowd  style=utl_ymrlan100  /* ut_pdflan100 */;
col title;
define title / display ' ';
run;
quit;

%Mend Tut_Sly;


%macro utl_boxpdf2ppt(inp=&outpdf001,out=&outppt001)/des="www.boxoft.con pdf to ppt";
  data _null_;
    cmd=catt("&pdf2ppt",' "',"&inp", '"',' "',"&out",'"');
    put cmd;
    call system(cmd);
  run;
%mend utl_boxpdf2ppt;

%MACRO greenbar ;
   DEFINE _row / order COMPUTED NOPRINT ;
   COMPUTE _row;
      nobs+1;
      _row = nobs;
      IF (MOD( _row,2 )=0) THEN
         CALL DEFINE( _ROW_,'STYLE',"STYLE={BACKGROUND=graydd}" );
   ENDCOMP;
%MEND greenbar;

%macro codebegin;
  options orientation=landscape ls=96;
  ods pdf;
  ods pdf bookmarkgen=off;
  data _null_;
   input;
   if _infile_ ne '';
   _infile_=trim(_infile_);
   file print;
   put _infile_;
   call execute(_infile_);
%mend codebegin;

%macro pdfbeg(rules=all,frame=box);
    %utlnopts;
    options orientation=landscape validvarname=v7;
    ods listing close;
    ods pdf close;
    ods path work.templat(update) sasuser.templat(update) sashelp.tmplmst(read);
    %utlfkil(&outpdf01.);
    ods noptitle;
    ods escapechar='^';
    ods listing close;
    ods graphics on / width=10in  height=7in ;
    ods pdf file="&wevoutpdf" style=utl_ymrlan100 pdftoc=1 bookmarkgen=on bookmarklist=show ;
    %utlopts;
%mend pdfbeg;

%macro pdfend;
   ods graphics off;
   ods pdf close;
   ods listing;
   %utlopts;
%mend pdfend;

 ***   *****    *    ****   *****
*   *    *     * *   *   *    *
 *       *    *   *  *   *    *
  *      *    *****  ****     *
   *     *    *   *  * *      *
*   *    *    *   *  *  *     *
 ***     *    *   *  *   *    *;


* common slide properties;
%let z=%str(                  );
%let b=%str(font_weight=bold);
%let c=%str(font=("Courier New"));
%let w=%str(cellwidth=100pct);

/*
* because I allow macro triggers
use these when you do not want a trigger in a slide.
Use double quotes when possible
` to single
| to ,
` to single quote
~ to semi colon
# to percent sign
@ to ambersand
*/

title;
footnote;

%utl_ymrlan100
    (
      style=utl_ymrlan100
      ,frame=void
      ,TitleFont=15pt
      ,docfont=14pt
      ,fixedfont=12pt
      ,rules=none
      ,bottommargin=.25in
      ,topmargin=.25in
      ,rightmargin=.25in
      ,leftmargin=.25in
      ,cellheight=15pt
      ,cellpadding = 2pt
      ,cellspacing = .2pt
      ,borderwidth = .2pt
    );

%pdfbeg;

%Tut_Sly
   (
    stop=7
    ,L6 ='^S={font_size=35pt just=c &w}Example of SAS Weave'
    ,L7 ='^S={font_size=35pt just=c &w}Documenting your programs'
   );

%Tut_Sly
   (
    stop=7
    ,L7 ='^S={font_size=25pt just=c &w}Documenting code and Outputs'
   );

%Tut_Sly
   (
    stop=7
    ,L6 ='^S={font_size=25pt just=c &w}Class Height and Weight by Sex'
    ,L7 ='^S={font_size=25pt just=c &w}Two different Y axes'
   );


%codebegin;
cards4;
ods pdf bookmarkgen=off;
proc report data=sashelp.class  split="/" nocenter missing ;
column  ("Sashelp class data set for plot" name sex age height weight) _row;

define  name / display format= $8. width=8     spacing=2   left "Name" ;
define  sex / display format= $1. width=1     spacing=2   center "Gender" ;
define  age / sum format= best9. width=9     spacing=2   center "Age" ;
define  height / sum format= best9. width=9     spacing=2   center "Height" ;
define  weight / sum format= best9. width=9     spacing=2   center "Weight" ;
%greenbar;
run;quit;
;;;;
run;quit;

%codebegin cards4;
ods pdf bookmarkgen=off;
proc template;
  define statgraph plot;
  begingraph;
  entrytitle "Series plot with addtional Y2AXIS" ;
    layout datapanel classvars=(sex) / columns=2 rows=1
       headerlabelattrs=(color=red weight=bold)
       headeropaque=false;
       layout prototype;
          seriesplot x=age y=height / group=dose display=all name="height"
             lineattrs=(color=cx5DAF5D)
             markerattrs=(symbol=squarefilled color=cx5DAF5D);
          seriesplot x=age y=weight / group=dose
             yaxis=y2 display=all name="weight"
             lineattrs=(color=cx4B50AA pattern=2)
             markerattrs=(symbol=circlefilled color=cx4B50AA);
       endlayout;
      sidebar / align=bottom;
         discretelegend "height" "weight" / ;
      endsidebar;
    endlayout;
  endgraph;
  end;
  define style noheaderborder;
     parent = styles.default;
     class graphborderlines / contrastcolor=white;
     class graphbackground / color=white ;
  end;

proc sort data=sashelp.class out=class;
by age;
run;quit;
proc sgrender data=class template=plot;
run;quit;
;;;;
run;quit;


%Tut_Sly
   (
    stop=7
    ,L6 ='^S={font_size=25pt just=c &w}Class Height and Weight by Sex'
    ,L7 ='^S={font_size=25pt just=c &w}Four plots on one page'
   );

%codebegin;
cards4;
ods pdf bookmarkgen=off;
proc template;
 define statgraph panel;
 begingraph;
 entrytitle "paneled display ";
   layout lattice / rows = 2 columns = 2 rowgutter = 10
      columngutter = 10;
     layout overlay; scatterplot y = weight x = height;
        regressionplot y = weight x = height;
      endlayout;
      layout overlay / xaxisopts = (label = "weight");
         histogram weight;
      endlayout;
      layout overlay / yaxisopts = (label = "height");
         boxplot y = height;
      endlayout;
      layout overlay; scatterplot y = weight x = height /
        group = sex name = "scat";
        discretelegend  "scat" /
          location=inside autoalign=(topleft) across=1;
      endlayout;
   endlayout;
 endgraph;
end;
run;quit;
proc sgrender data = sashelp.class template = panel;
run;quit;
;;;;
run;quit;

%pdfend;


* convert pdf slides to ppt slides;
%utl_boxpdf2ppt(inp=&wevoutpdf,out=&wevoutppt);


/* adding your theme;
Win 7 32 bit Powerpoint 2010

Once you have created your slides using ODS PDF and converted them to
a ppt using the free BOXOFT PDF2PPT, you probably want to
add a powerpoint theme. Make sure you used a SAS template that
provides whitespace for the banner and slide number footnote.

If you have a company custom template click on it and
select all the slides, shift first slide then last slide,
Then go to design tab then background slides then apply.
All slides will now have then them behind the SAS foreground
bitmaps.

The following VBA macro will make the SAS bitmaps transparent
and the company theme will appear on each slide.

Note SAS does not provide the ability for a them using
the template. Since SAS titles always appear at the top of a
page you can imbed an image in title1. However I wanted the
company powerpoint theme.

Thanks to the Powerpoint Expert Steve Rindsberg for
the macro below

Sub SetXparency()


Dim oSl As Slide

For Each oSl In ActivePresentation.Slides

    With oSl.Shapes(1)

        With .PictureFormat

            .TransparentBackground = msoTrue

            .TransparencyColor = RGB(255, 255, 255)

        End With

    End With

Next

End Sub

*/
