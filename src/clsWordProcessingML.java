import java.io.File;  
import java.math.BigInteger;

import javax.xml.bind.JAXBElement;
import javax.xml.namespace.QName;

import org.docx4j.UnitsOfMeasurement;
import org.docx4j.jaxb.Context;
import org.docx4j.model.properties.table.tc.AbstractTcProperty;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;  
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTShd;
import org.docx4j.wml.CTSimpleField;
import org.docx4j.wml.CTTblPrBase.TblStyle;
import org.docx4j.wml.CTView;
import org.docx4j.wml.FldChar;
import org.docx4j.wml.FooterReference;
import org.docx4j.wml.HdrFtrRef;
import org.docx4j.wml.HeaderReference;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.STFldCharType;
import org.docx4j.wml.STView;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.TblWidth;
import org.w3c.dom.css.CSSPrimitiveValue;
import org.w3c.dom.css.CSSValue;

public class clsWordProcessingML {
	//for main document setting count
	final static int c_tr=12;
	final static int c_tc=6;
	final static int c_p=6;
	final static int c_r=6;
	final static int c_t=6;
	final static int c_tcpr=6;
	final static int c_rpr=6;
	final static int c_color=6;
	final static int c_shd=6;
	final static int c_hpsmeasure=6;
	
	
	//for header document setting count
	final static int ch_tr=2;
	final static int ch_tc=6;
	final static int ch_p=6;
	final static int ch_r=6;
	final static int ch_t=6;
	final static int ch_tcpr=6;
	final static int ch_rpr=6;
	final static int ch_color=6;
	final static int ch_shd=6;
	final static int ch_hpsmeasure=6;
	
	
	//for header document setting count
	final static int cf_tr=1;
	final static int cf_tc=3;
	final static int cf_p=3;
	final static int cf_r=3;
	final static int cf_t=3;
	final static int cf_tcpr=3;
	final static int cf_rpr=3;
	final static int cf_color=3;
	final static int cf_shd=3;
	final static int cf_hpsmeasure=3;
		
	
	//docx4j main object
	WordprocessingMLPackage wordMLPackage;
	static org.docx4j.wml.ObjectFactory factory;
	static SectPr sectPr;
	static CTSimpleField ctsimplefield=new CTSimpleField();
	static FldChar begin_fldchar=new FldChar();
	static FldChar end_fldchar=new FldChar();
	
	//docx4j main document struct object
	static org.docx4j.wml.Tbl tbl=new org.docx4j.wml.Tbl();
	static org.docx4j.wml.Tr tr[]=new org.docx4j.wml.Tr[c_tr];
	static org.docx4j.wml.Tc tc[]=new org.docx4j.wml.Tc[c_tc];
	static org.docx4j.wml.P p[]=new org.docx4j.wml.P[c_p];
	static org.docx4j.wml.R r[]=new org.docx4j.wml.R[c_r];
	static org.docx4j.wml.Text t[]=new org.docx4j.wml.Text[c_t];
	
	static org.docx4j.wml.TblPr tblpr=new org.docx4j.wml.TblPr();
	static org.docx4j.wml.TcPr tcpr[]=new org.docx4j.wml.TcPr[c_tcpr];
	static org.docx4j.wml.RPr rpr[]=new org.docx4j.wml.RPr[c_rpr];
	
	static TblWidth tblw=new TblWidth();
	static org.docx4j.wml.TblBorders tblborders=new org.docx4j.wml.TblBorders();
	static org.docx4j.wml.CTBorder ctborder=new org.docx4j.wml.CTBorder();
	static org.docx4j.wml.Color color[]=new org.docx4j.wml.Color[c_color];
	static org.docx4j.wml.CTShd shd[]=new org.docx4j.wml.CTShd[c_shd];
	static org.docx4j.wml.HpsMeasure hpsmeasure[]=new org.docx4j.wml.HpsMeasure[c_hpsmeasure];
	
	
	//docx4j header document struct object
	static HeaderPart headerPart;
	static Relationship header_relationship;
	static HeaderReference headerReference;
	
	static org.docx4j.wml.Hdr header_hdr = new org.docx4j.wml.Hdr();

	static org.docx4j.wml.Tbl header_tbl=new org.docx4j.wml.Tbl();
	static org.docx4j.wml.Tr header_tr[]=new org.docx4j.wml.Tr[ch_tr];
	static org.docx4j.wml.Tc header_tc[]=new org.docx4j.wml.Tc[ch_tc];
	static org.docx4j.wml.P header_p[] = new org.docx4j.wml.P[ch_p];
	static org.docx4j.wml.R header_r[] = new org.docx4j.wml.R[ch_r];
	static org.docx4j.wml.Text header_t[] = new org.docx4j.wml.Text[ch_t];
	
	static org.docx4j.wml.TblPr header_tblpr=new org.docx4j.wml.TblPr();
	static org.docx4j.wml.RPr header_rpr[] = new org.docx4j.wml.RPr[ch_rpr];
	
	static TblWidth header_tblw=new TblWidth();
	static org.docx4j.wml.HpsMeasure header_hpsmeasure[] = new org.docx4j.wml.HpsMeasure[ch_hpsmeasure];
	static org.docx4j.wml.Color header_color[]=new org.docx4j.wml.Color[ch_color];
	static org.docx4j.wml.TblBorders header_tblborders=new org.docx4j.wml.TblBorders();
	static org.docx4j.wml.CTBorder header_ctborder=new org.docx4j.wml.CTBorder();
	
	//docx4j footer document struct object
	static FooterPart footerPart;
	static Relationship footer_relationship;
	static FooterReference footerReference;
	
	static org.docx4j.wml.Ftr footer_ftr = new org.docx4j.wml.Ftr();
	
	static org.docx4j.wml.Tbl footer_tbl=new org.docx4j.wml.Tbl();
	static org.docx4j.wml.Tr footer_tr[]=new org.docx4j.wml.Tr[cf_tr];
	static org.docx4j.wml.Tc footer_tc[]=new org.docx4j.wml.Tc[cf_tc];
	static org.docx4j.wml.P footer_p[] = new org.docx4j.wml.P[cf_p];
	static org.docx4j.wml.R footer_r[] = new org.docx4j.wml.R[cf_r];
	static org.docx4j.wml.Text footer_t[] = new org.docx4j.wml.Text[cf_t];
	
	static org.docx4j.wml.TblPr footer_tblpr=new org.docx4j.wml.TblPr();
	static org.docx4j.wml.RPr footer_rpr[] = new org.docx4j.wml.RPr[cf_rpr];
	
	static TblWidth footer_tblw=new TblWidth();
	static org.docx4j.wml.HpsMeasure footer_hpsmeasure[] = new org.docx4j.wml.HpsMeasure[cf_hpsmeasure];
	static org.docx4j.wml.Color footer_color[]=new org.docx4j.wml.Color[cf_color];
	
	//docx4j other struct object
	static org.docx4j.wml.BooleanDefaultTrue bdt=new org.docx4j.wml.BooleanDefaultTrue();
	static org.docx4j.wml.BooleanDefaultFalse bdf=new org.docx4j.wml.BooleanDefaultFalse();
	
	
	//initalize docx4j main object
    protected void initalize() throws InvalidFormatException{
        //create word file
        wordMLPackage=WordprocessingMLPackage.createPackage();  
        //MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
        //createHeaderPart(wordMLPackage);
        
        //open old data
        //WordprocessingMLPackage wordMLPackage=WordprocessingMLPackage.load(new java.io.File(inputfilepath));
        //MainDocumentPart documentPart=wordMLPackage.getMainDocumentPart();
        //org.docx4j.wml.Document wmlDocumentEl=(org.docx4j.wml.Document)documentPart.getJaxbElement();
        //Body body=wmlDocumentEl.getBody();
        
        factory=new org.docx4j.wml.ObjectFactory();
    }
    
    
    protected void factory_create(){
    	int i;
    	
    	tbl=factory.createTbl();
    	for(i=0;i<c_tr;i++){
    		tr[i]=factory.createTr();
    	}
    	for(i=0;i<c_tc;i++){
    		tc[i]=factory.createTc();
    	}
    	for(i=0;i<c_p;i++){
    		p[i]=factory.createP();
    	}
    	for(i=0;i<c_r;i++){
    		r[i]=factory.createR();
    	}
    	for(i=0;i<c_t;i++){
    		t[i]=factory.createText();
    	}
        
    	tblpr=factory.createTblPr();
    	for(i=0;i<c_tcpr;i++){
    		tcpr[i]=factory.createTcPr();
    	}
    	for(i=0;i<c_rpr;i++){
    		rpr[i]=factory.createRPr();
    	}
        
    	tblborders=factory.createTblBorders();
    	ctborder=factory.createCTBorder();
    	for(i=0;i<c_color;i++){
    		color[i]=factory.createColor();
    	}
    	for(i=0;i<c_shd;i++){
    		shd[i]=factory.createCTShd();
    	}
    	for(i=0;i<c_hpsmeasure;i++){
    		hpsmeasure[i]=factory.createHpsMeasure();
    	}
    }
    
    
    protected void title(){
    	//row 1, col 1
    	hpsmeasure[0].setVal(new java.math.BigInteger("20"));
        rpr[1].setSz(hpsmeasure[0]);
    	color[0].setVal("00FF00");
        rpr[0].setColor(color[0]);
        r[0].setRPr(rpr[0]);        
        
    	t[0].setValue("列印人員:$name$($nowdate$)");
    	r[0].getContent().add(t[0]);
    	p[0].getContent().add(r[0]);
    	tc[0].getContent().add(p[0]);
    	tr[0].getContent().add(tc[0]);
    	
    	//row 1, col 2
    	hpsmeasure[1].setVal(new java.math.BigInteger("30"));
        rpr[1].setSz(hpsmeasure[1]);
    	rpr[1].setB(bdt);
        r[1].setRPr(rpr[1]);
        
    	t[1].setValue("國泰敦南健檢中心");
    	r[1].getContent().add(t[1]);
    	p[1].getContent().add(r[1]);
    	tc[1].getContent().add(p[1]);
    	tr[0].getContent().add(tc[1]);
    	
    	//row 1, col 3
    	hpsmeasure[2].setVal(new java.math.BigInteger("40"));
        rpr[2].setSz(hpsmeasure[2]);
    	rpr[2].setB(bdt);
        r[2].setRPr(rpr[2]);
        
        t[2].setValue("*$chartno$*");
    	r[2].getContent().add(t[2]);
    	p[2].getContent().add(r[2]);
    	tc[2].getContent().add(p[2]);
    	tr[0].getContent().add(tc[2]);
    	
    	//row 2, col 1
    	t[3].setValue("");
    	r[3].getContent().add(t[3]);
    	p[3].getContent().add(r[3]);
    	tc[3].getContent().add(p[3]);
    	tr[1].getContent().add(tc[3]);
    	
    	//row 2, col 2
    	hpsmeasure[4].setVal(new java.math.BigInteger("30"));
        rpr[1].setSz(hpsmeasure[4]);
    	rpr[4].setB(bdt);
        r[4].setRPr(rpr[4]);
        
    	t[4].setValue("敦南健檢健檢報告(院內)");
    	r[4].getContent().add(t[4]);
    	p[4].getContent().add(r[4]);
    	tc[4].getContent().add(p[4]);
    	tr[1].getContent().add(tc[4]);
    	
    	//row 2, col 3
    	t[5].setValue("");
    	r[5].getContent().add(t[5]);
    	p[5].getContent().add(r[5]);
    	tc[5].getContent().add(p[5]);
    	tr[1].getContent().add(tc[5]);
    	
    	
        tbl.getContent().add(tr[0]);
        tbl.getContent().add(tr[1]);

        
        wordMLPackage.getMainDocumentPart().addObject(tbl);
    }
    
    protected void unsolved(){
    	//row 1,col 1
    	t[0].setValue("");
    	r[0].getContent().add(t[0]);
    	p[0].getContent().add(r[0]);
    	tc[0].getContent().add(p[0]);
    	tr[0].getContent().add(tc[0]);
    	
    	
    	tbl.getContent().add(tr[0]);
    	
    	
    	wordMLPackage.getMainDocumentPart().addObject(tbl);
    }
    
    protected void profile(){
    	//row 1, col 1
    	color[0].setVal("732303");
        rpr[0].setColor(color[0]);
        r[0].setRPr(rpr[0]);      
    	shd[0].setFill(UnitsOfMeasurement.rgbTripleToHex(230, 230, 230));
        tcpr[0].setShd(shd[0]);
        tc[0].setTcPr(tcpr[0]);
        
    	t[0].setValue("個人資料(PERSONAL INFORMATION)");
    	r[0].getContent().add(t[0]);
    	p[0].getContent().add(r[0]);
    	tc[0].getContent().add(p[0]);
    	tr[0].getContent().add(tc[0]);
    	
        //row 2, col 1
    	t[1].setValue("姓名(Name)");
    	r[1].getContent().add(t[1]);
    	p[1].getContent().add(r[1]);
    	tc[1].getContent().add(p[1]);
    	tr[1].getContent().add(tc[1]);
    	
    	//row 2, col 2
    	shd[2].setFill(UnitsOfMeasurement.rgbTripleToHex(230, 230, 230));
        tcpr[2].setShd(shd[2]);
        tc[2].setTcPr(tcpr[2]);
        
    	t[2].setValue("$chartname$");
    	r[2].getContent().add(t[2]);
    	p[2].getContent().add(r[2]);
    	tc[2].getContent().add(p[2]);
    	tr[1].getContent().add(tc[2]);

    	
    	tbl.getContent().add(tr[0]);
    	tbl.getContent().add(tr[1]);
    	
    	
    	wordMLPackage.getMainDocumentPart().addObject(tbl);
    }
    
    protected void content(){
    	//row 1, col 1
    	color[0].setVal("732303");
        rpr[0].setColor(color[0]);
        r[0].setRPr(rpr[0]);      
    	shd[0].setFill(UnitsOfMeasurement.rgbTripleToHex(230, 230, 230));
        tcpr[0].setShd(shd[0]);
        tc[0].setTcPr(tcpr[0]);

    	t[0].setValue("目錄(CONTENTS)");
    	r[0].getContent().add(t[0]);
    	p[0].getContent().add(r[0]);
    	tc[0].getContent().add(p[0]);
    	tr[0].getContent().add(tc[0]);
        
    	
    	tbl.getContent().add(tr[0]);
    	
    	
    	wordMLPackage.getMainDocumentPart().addObject(tbl);
    }
    
    
    protected void TOC_factory_create(){
    	begin_fldchar = factory.createFldChar();
    	end_fldchar = factory.createFldChar();
    }
    
    protected void TOC(){
    	//setting TOC
        begin_fldchar.setFldCharType(STFldCharType.BEGIN);
        begin_fldchar.setDirty(true);
        r[0].getContent().add(getWrappedFldChar(begin_fldchar));

        t[0].setSpace("preserve");
        t[0].setValue("TOC \\o \"1-1\" \\h \\z \\u \\h");
        r[0].getContent().add(factory.createRInstrText(t[0]));
        
        end_fldchar.setFldCharType(STFldCharType.END);
        r[0].getContent().add(getWrappedFldChar(end_fldchar));
        
        p[0].getContent().add(r[0]);
        
        
        //add to main document part
        wordMLPackage.getMainDocumentPart().addObject(p[0]);
       
        
        //setup TOC style
        wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "Hello 1");
        wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "Hello 2");
    }
    
    
    
    protected void header_factory_create() throws InvalidFormatException {
    	int i;
    	
    	headerPart = new HeaderPart();
    	
   		headerReference = factory.createHeaderReference();
   		
    	header_hdr = factory.createHdr();

    	
    	header_tbl=factory.createTbl();
    	for(i=0;i<ch_tr;i++){
    		header_tr[i]=factory.createTr();
    	}
    	for(i=0;i<ch_tc;i++){
    		header_tc[i]=factory.createTc();
    	}
    	for(i=0;i<ch_p;i++){
    		header_p[i] = factory.createP();
    	}
    	for(i=0;i<ch_r;i++){
    		header_r[i] = factory.createR();
    	}
    	for(i=0;i<ch_t;i++){
    		header_t[i] = factory.createText();
    	}
    	
    	for(i=0;i<ch_hpsmeasure;i++){
    		header_hpsmeasure[i]=factory.createHpsMeasure();
    	}
    	for(i=0;i<ch_rpr;i++){
    		header_rpr[i]=new org.docx4j.wml.RPr();
    	}
    	for(i=0;i<ch_color;i++){
    		header_color[i]=new org.docx4j.wml.Color();
    	}
    }
    
    
    protected void header() throws InvalidFormatException{    	
    	//header row 1, col 1
    	header_hpsmeasure[0].setVal(new java.math.BigInteger("20"));
        header_rpr[0].setSz(header_hpsmeasure[0]);
        header_r[0].setRPr(header_rpr[0]);  
    	
    	header_t[0].setValue("健管號碼:$chartno$");
    	header_r[0].getContent().add(header_t[0]);
    	header_p[0].getContent().add(header_r[0]);
    	header_tc[0].getContent().add(header_p[0]);
    	header_tr[0].getContent().add(header_tc[0]);
    	
    	//header row 1, col 2
    	header_hpsmeasure[1].setVal(new java.math.BigInteger("20"));
        header_rpr[1].setSz(header_hpsmeasure[0]);
    	header_color[1].setVal("00FF00");
        header_rpr[1].setColor(header_color[1]);
        header_r[1].setRPr(header_rpr[1]);  
    	
    	header_t[1].setValue("國泰健康管理");
    	header_r[1].getContent().add(header_t[1]);
    	header_p[1].getContent().add(header_r[1]);
    	header_tc[1].getContent().add(header_p[1]);
    	header_tr[0].getContent().add(header_tc[1]);
    	
    	//header row 1, col 3
    	header_t[2].setValue("檢查日期：$examdate$");
    	header_r[2].getContent().add(header_t[2]);
    	header_p[2].getContent().add(header_r[2]);
    	header_tc[2].getContent().add(header_p[2]);
    	header_tr[0].getContent().add(header_tc[2]);
   
    	
    	//setting table
    	header_tblw.setW(BigInteger.valueOf(8700));
    	header_tblw.setType("xda");
    	header_tblpr.setTblW(header_tblw);

    	header_ctborder.setVal(STBorder.CHECKED_BAR_BLACK);
    	header_ctborder.setColor("00FF00");
    	header_ctborder.setSz(BigInteger.valueOf(30));
    	header_tblborders.setBottom(header_ctborder);
    	header_tblpr.setTblBorders(header_tblborders);
    	
    	header_tbl.setTblPr(header_tblpr);

    	header_tbl.getContent().add(header_tr[0]);
    	
    	
    	header_hdr.getContent().add(header_tbl);
    	
    	
    	headerPart.setJaxbElement(header_hdr);
   		header_relationship = wordMLPackage.getMainDocumentPart().addTargetPart(headerPart);
 		
   		
   		headerReference.setId(header_relationship.getId());
   		headerReference.setType(HdrFtrRef.DEFAULT);
   		sectPr.getEGHdrFtrReferences().add(headerReference);
    }
    
    
    
    protected void footer_factory_create() throws InvalidFormatException{
    	int i;
    	
    	footerPart=new FooterPart();
    	
   		footerReference = factory.createFooterReference();
   		
    	footer_ftr = factory.createFtr();

    	
    	footer_tbl=factory.createTbl();
    	for(i=0;i<cf_tr;i++){
    		footer_tr[i]=factory.createTr();
    	}
    	for(i=0;i<cf_tc;i++){
    		footer_tc[i]=factory.createTc();
    	}
    	for(i=0;i<cf_p;i++){
    		footer_p[i] = factory.createP();
    	}
    	for(i=0;i<cf_r;i++){
    		footer_r[i] = factory.createR();
    	}
    	for(i=0;i<cf_t;i++){
    		footer_t[i] = factory.createText();
    	}
    	
    	for(i=0;i<cf_hpsmeasure;i++){
    		footer_hpsmeasure[i]=factory.createHpsMeasure();
    	}
    	for(i=0;i<cf_rpr;i++){
    		footer_rpr[i]=new org.docx4j.wml.RPr();
    	}
    	for(i=0;i<cf_color;i++){
    		footer_color[i]=new org.docx4j.wml.Color();
    	}
    }
    
    protected void footer() throws InvalidFormatException{    	
    	//footer row 1, col 1
    	footer_hpsmeasure[0].setVal(new java.math.BigInteger("20"));
        footer_rpr[0].setSz(footer_hpsmeasure[0]);
        footer_color[0].setVal("00FF00");
        footer_rpr[0].setColor(footer_color[0]);
        footer_r[0].setRPr(footer_rpr[0]);  
    	
    	footer_t[0].setValue("高額保戶免費體檢");
    	footer_r[0].getContent().add(footer_t[0]);
    	footer_p[0].getContent().add(footer_r[0]);
    	footer_tc[0].getContent().add(footer_p[0]);
    	footer_tr[0].getContent().add(footer_tc[0]);
    	
    	//footer row 1, col 2
    	footer_hpsmeasure[1].setVal(new java.math.BigInteger("20"));
        footer_rpr[1].setSz(footer_hpsmeasure[0]);
        footer_r[1].setRPr(footer_rpr[1]);  
    	
        ctsimplefield.setInstr(" PAGE \\* MERGEFORMAT ");
        footer_r[1].getContent().add(ctsimplefield);
    	footer_p[1].getContent().add(footer_r[1]);
    	footer_tc[1].getContent().add(footer_p[1]);
    	footer_tr[0].getContent().add(footer_tc[1]);
    	
    	//footer row 1, col 3
    	footer_hpsmeasure[2].setVal(new java.math.BigInteger("20"));
        footer_rpr[2].setSz(footer_hpsmeasure[2]);
        footer_color[2].setVal("00FF00");
        footer_rpr[2].setColor(footer_color[2]);
        footer_r[2].setRPr(footer_rpr[2]);  
    	
    	footer_t[2].setValue("國泰健康管理．守護永續不息");
    	footer_r[2].getContent().add(footer_t[2]);
    	footer_p[2].getContent().add(footer_r[2]);
    	footer_tc[2].getContent().add(footer_p[2]);
    	footer_tr[0].getContent().add(footer_tc[2]);
    	
    	
    	footer_tbl.getContent().add(footer_tr[0]);
    	
    	
    	footer_ftr.getContent().add(footer_tbl);
    	
    	
    	footerPart.setJaxbElement(footer_ftr);
   		footer_relationship = wordMLPackage.getMainDocumentPart().addTargetPart(footerPart);
 		
   		
   		footerReference.setId(footer_relationship.getId());
   		footerReference.setType(HdrFtrRef.DEFAULT);
   		sectPr.getEGHdrFtrReferences().add(footerReference);
    }
    
    
    protected void sectPr_factory_create(){
    	sectPr = factory.createSectPr();
    }
    protected void set_sectPr(){
    	wordMLPackage.getMainDocumentPart().addObject(sectPr);	
    }
    
    protected void ctsimplefield_factory_create(){
    	ctsimplefield = factory.createCTSimpleField();
    }
    
    public static JAXBElement getWrappedFldChar(FldChar fldchar) {
        return new JAXBElement(new QName(Namespaces.NS_WORD12, "fldChar"), FldChar.class, fldchar);   
    }
}


//add by string
//wordMLPackage.getMainDocumentPart().addObject(org.docx4j.XmlUtils.unmarshalString(str));
