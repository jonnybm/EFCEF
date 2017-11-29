import java.io.BufferedInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.sql.Date;

import javax.faces.bean.ManagedBean;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;

import com.sun.prism.paint.Color;
import com.sun.xml.internal.ws.util.StringUtils;
 


@ManagedBean
public class EFCEF {
	
	private static GetSetCEF getSetCEF = new GetSetCEF();
	private static EficienciaFinanceiraBB EficienciaFinanceiraBB = new EficienciaFinanceiraBB();
	private static String banco = "CEF";  // BB ou CEF
	
//  	static String caminho = "/Volumes/HD/BackupJonny/Projetos/EficienciaFinanceira/PDFBB/";
//  	static String excelBB = "/Volumes/HD/BackupJonny/Projetos/EficienciaFinanceira/ExcelBB/Arcos31.07.2017.xlsx";

  	static String caminho = "";
  	static String excelBB = "";
  	static String ComboMes = "";
  	static String ComboAno = "";
	
	
  	static ArrayList<String> arrayPublicoCJExisteUnica = new ArrayList<String> ();
  	
  	
  	static File fo = new File(excelBB);

	//public static void main(String[] args) throws IOException {	
	public static void init() throws IOException {
		fo = new File(excelBB);
		visualizarArquivos();
		
	}

	
	  public static void visualizarArquivos() throws IOException {
		  


		  	String arquivoPDF = "";	
		  	String ret = "";
		  	String excluidos  = "";
		  
		  	File file = new File(caminho);
			File afile[] = file.listFiles();
			int i = 0;
			
			for (int j = afile.length; i < j; i++) {
				File arquivos = afile[i];
				arquivoPDF = caminho+arquivos.getName();
				//System.out.println(arquivos.getName());
				
				if(arquivos.getName().indexOf ("pdf") >= 0) {
				
					//Recebe e Le Texto do PDF
					String texto = extraiTextoDoPDF(arquivoPDF);
					
					//Recebe o Conteudo do PDF e coloca em Array com quebra de Linha
					String linhas[] = texto.split("\n");
					
			        					        			
					//Trata o PDF lendo linha a linha do array
					//String textoret = trataPDF(linhas);
					 ret = trataPDF(linhas);
					
					 System.out.println("PASSo 5 {"+ ret+"}");
					
					 //System.out.println("--> "+textoret);
										 
					  if(arquivoPDF.indexOf ("DS_Store") <= 0 ) //Se arquivo for diferente de arquivo de sistema que nao precisa ser analizado ou se o valor é zero mas nao igual ao mes de consulta
					  {
						if(!ret.equals("ValorZeroNaoMesDeReferenciaContulta"))
						{ 
							lerExcel("a");
						  
							getSetCEF.setPorcentagem((i*100)/afile.length);
							System.out.println(" STATUS: ["+getSetCEF.getPorcentagem()+" %]");
						}
						else
						{
							excluidos = caminho+"/EXCLUIDOS";
							diretorio(excluidos);
							
							// -------- FAZER AQUI ------- ///
				
							// ----2 SE entrou neste if pod jogar o PDF (arquivoPDF) para a pasta EXCLUIDOS
							// --  3 
						}	
					  }
				}
			
				

			}
			
			
			getSetCEF.setFimArquivo("FINALIZADO COM SUCESSO");
			System.out.println("FINALIZADO COM SUCESSO");
		}
	
  

    		//EXTRAI OS TEXTOS DE DENTRO DO PDF
		  public static String extraiTextoDoPDF(String caminho) {
			  
			  if(caminho.indexOf ("DS_Store") >= 0) {
				  return "NAOePDF";
			  }
			  
			    PDDocument pdfDocument = null;
			    try {
			      pdfDocument = PDDocument.load(caminho);
			      PDFTextStripper stripper = new PDFTextStripper();
			      String texto = stripper.getText(pdfDocument);
			      return texto;
			    } catch (IOException e) {
			      throw new RuntimeException(e);
			    } finally {
			      if (pdfDocument != null) try {
			        pdfDocument.close();
			      } catch (IOException e) {
			        throw new RuntimeException(e);
			      }
			    }
		  }
			  
		  
		  //RECEBE O CONTEUDO DO EXCEL PARA TRATAMENTO DAS INFORMACOES
		  public static String trataPDF(String[] linhas) 
		  {			  
			  String ret = "";
			  
			  //System.out.println("PASSO 1");

				  if(banco.equals("CEF")) // Banco do Brasil
				  {
					  for (int i = 0; i < linhas.length; i++) 
					  {
						  
						  ret = linhas[i];
						 // System.out.println("CONTEUDO :>"+ret + " ["+i+"]");
						  
						  //PEGANDO A CONTA JUDICIAL
						  if(i == 20 ) //PEGANDO CONTA JUDICIAL
				          {
				           		ret = linhas[i];
				           		//ret =  trataSujeira(ret.trim());
				            		
				            		//SETANDO A CONTA JUDICIAL
				            		getSetCEF.setContaJuridica(ret);    
				            		
				            		//System.out.println("CONTA JU : ["+ret+"]");
				          }
						  //LINHA 8 PEGA O NOME AUTOR
						  if(i == 25)
				          {
							  	ret = "";
			        	  			ret = linhas[i];	
			        	  			
				        	  		ret = ret.replace("1", "");
				        	  		ret = ret.replace("2", "");
				        	  		ret = ret.replace("3", "");
				        	  		ret = ret.replace("4", "");
				        	  		ret = ret.replace("5", "");
				        	  		ret = ret.replace("6", "");
				        	  		ret = ret.replace("7", "");
				        	  		ret = ret.replace("8", "");
				        	  		ret = ret.replace("9", "");
				        	  		ret = ret.replace("0", "");
				        	  		ret = ret.replace(".", "");
				        	  		ret =  trataSujeira(ret);
				        	  		ret =  toTitledCase(ret);
				            		
				            		//SETANDO O AUTOR
				            		getSetCEF.setAutor(ret);
				            		//System.out.println("AUTOR :>"+ret);
				          }
						  
				          //PEGANDO O VALOR
				          if(linhas.length-2 == i)//é o penultimo  Registro?
				          {
				        	  			
				        	  		if(!ret.matches("[0-9]"))//contem apenas numero?
				        	        {
				        	  			
					        	  		String[] Valor = null;
					        	  		Valor =  linhas[i].substring(5, linhas[i].length()-1).split(" ");// Pega a linha e Tira a Data que contem incialmente na linha faz Splito para pegar o segundo valor da Linha
					        	  		
					        	  		//DATA ANO CONSULTA
					        	  		ret = "";
					        	  		ret = linhas[i].substring(6, 10);
					        	  		getSetCEF.setDataAnoConsulta(ret.trim());
					        	  		System.out.println("ANO EXTRATO :>"+ret);

					        	  		
					        	  		
					        	  		//DATA MES CONSULTA
					        	  		ret = "";
					        	  		ret = linhas[i].substring(3, 5);
					        	  		getSetCEF.setDataMesConsulta(ret.trim());
					        	  		
					        	  		
				        	  			//VAI RODAR as ultima 4linhas de tras pra frente
				        	  			for(int i1 = 1; i1 < 6; ++i1){
				        	  				
				        	  				
				        	  				//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
				        	  				//>>>>>>>> PERGUNTAR O ANO TAMBEM <<<<<<<<
				        	  				//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>				        	  				
				        	  				if( getSetCEF.getDataMesConsulta().equals(ComboMes) && getSetCEF.getDataAnoConsulta().equals(ComboAno))  
				        	  				{
				        	  					
				        	  					//System.out.println("DATA IGUAIS VAI PEGAR VALOR :>"+ComboMes);
				        	  												        	  		
							        	  		//VALOR
				        	  					ret = "";
							        	  		ret = Valor[2];
							        	  		ret =  trataSujeira(ret);
							        	  		getSetCEF.setValorAtualizado(ret.trim());
				        	  					
				        	  					break;
				        	  				}
				        	  				else 
				        	  				{
				        	  					
				        	  					//System.out.println("VALOR I :>"+i1);
//				        	  					System.out.println("VALOR LINHA :>"+ i1);
				        	  					
				        	  					if(i1 == 3 ) //ultima contagem do For???
				        	  					{	
				        	  					
				        	  						Valor =  linhas[i].substring(5, linhas[i].length()-1).split(" ");// Pega a linha e Tira a Data que contem incialmente na linha faz Splito para pegar o segundo valor da Linha
								        	  	

								        	  		//VALOR
				        	  						ret = "";
								        	  		ret = Valor[2];
								        	  		ret =  trataSujeira(ret);
								        	  		ret = ret.trim();
								        	 
								        	  		System.out.println("VALOR LINHA : ["+ret.trim()+"]");
								        	  		
								        	  		//Se VALOR FOR ZERO E NAO ESTIVER NO MES DE REFERENCIA DE CONSULTA IGNORAR	
								        	  		if(ret.equals("0,00"))
								        	  		{	
								        	  			
								        	  			ret = "ValorZeroNaoMesDeReferenciaContulta";
								        	  			
								        	  			return ret;
								        	  			
								        	  		}
								        	  		else
								        	  		{
								        	  			getSetCEF.setValorAtualizado(ret.trim());
								        	  		}
				        	  					}
				        	  					else
				        	  					{				        	  						
				        	  						Valor =  linhas[i-i1].substring(5, linhas[i-i1].length()-1).split(" ");// Pega a linha e Tira a Data que contem incialmente na linha faz Splito para pegar o segundo valor da Linha
								        	  							        	  					
								        	  		//DATA MES CONSULTA
				        	  						ret = "";
								        	  		ret = linhas[i-i1].substring(3, 5);
								        	  		
								        	  		//System.out.println("DATA CORReSPONDENTE :>"+ret);
								        	  		getSetCEF.setDataMesConsulta(ret.trim());
				        	  					}	
				        	  				}
				        	  			}
				        	        }
				        	  		else {
				        	  			ret =  ret; //se sim para nao dar erro envia apenas o nome
				        	  		}
				        	  		
				          }
						  
						  
						  
						  
						  if(i == 17) //LINHA 3 PEGA O PROCESSO
				          {
				            		ret = linhas[i];
				            		ret =  trataSujeira(ret);
				            		//SETANDO O PROCESSO
				            		getSetCEF.setProcesso(ret);
				            		
				            		System.out.println("PROCESSO :>"+ret.trim());
				          }
						  						  
						  if(i == 16) //LINHA 5 PEGA VARA , COMARCA e ESTADO
				          {
							 // System.out.println("PASSO 2");	
							  
							  ret = linhas[i];
				            		
				            	 	String[] parts = null;
				            	 	parts = ret.split("-");
				            	 					            	 	
				            	 	String[] partsVara = null;
				            	 	partsVara = ret.split("-");
				            	 	
				            	 //	System.out.println("PASSO 2.3");
				            		
				            	 	//VARA DO TRABALHO
				            	 	ret = partsVara[0];
				            	 	ret =  trataSujeira(ret);
				            	 	ret = ret.trim()+"ª VT";
				            	 	getSetCEF.setVara(ret);
				            		//System.out.println("VARA :>"+ret.trim());

				            	 	//System.out.println("PASSO 2.4");
				            	 	
				            	 	
				            	 	//COMARCA
				            		ret = parts[1];
				            		
				            		System.out.println("COMARCA E ESTADO :>"+ret.trim());
				            		
				            	 	parts = ret.split("/");
				            	 	
				            	 	
				            	 	ret = parts[0];
				            	 	ret =  toTitledCase(ret);
				            	 	getSetCEF.setComarca(ret);
				            		//System.out.println("COMARCA :>"+ret.trim());
				            	 	
				            	 	
				            		
				            	 	//ESTADO - Seta Estado
				            		
				            		if(parts.length == 1)
				            		{
				            			ret = " ";	
				            		}
				            		else
				            		{	
					            		ret = parts[1];
					            	 	getSetCEF.setEstado(ret);
				            		}

				          }

						  
						  if(i == 24) //LINHA 7 PEGA O RÉU E CNPJ
				          {
							  	//CNPJ
				            		ret = linhas[i];
				            		
				            		ret = ret.substring(0,18);
				            		getSetCEF.setCNPJ(ret);
				            		//System.out.println("CNPJ :>"+ret);
				            		
				            		
				            		//REU
				            		ret = linhas[i];
				            		ret =  trataSujeira(ret);
				            		
				        	  		ret = ret.replace("1", "");
				        	  		ret = ret.replace("2", "");
				        	  		ret = ret.replace("3", "");
				        	  		ret = ret.replace("4", "");
				        	  		ret = ret.replace("5", "");
				        	  		ret = ret.replace("6", "");
				        	  		ret = ret.replace("7", "");
				        	  		ret = ret.replace("8", "");
				        	  		ret = ret.replace("9", "");
				        	  		ret = ret.replace("0", "");
				        	  		ret = ret.replace(".", "");
				        	  		
				            		ret =  toTitledCase(ret);
				            		
				            		//SETANDO O AUTOR
				            		getSetCEF.setReu(ret);
				            		//System.out.println("REU :>"+ret);
				          }
		  
						  if(i == 28) //LINHA 9 PEGA O VALOR DO DEPOSITO INICIAL
				          {
				            	ret = linhas[i];
				            	
				            	
				            	
				            	String[] partsVara = null;
			            	 	partsVara = ret.split(" ");
			            		
			            	 	//DATA DEPOSITO
			            	 	ret = partsVara[0];
				            	getSetCEF.setDataDeposito(ret);
				            	//System.out.println("DATA DEPOSITO :>"+ret);
			            	 	
				            	
				            	
			            	 	//VARA DO TRABALHO
			            	 	ret = partsVara[1];
			            	 	ret =  trataSujeira(ret.trim());	
				            	//SETANDO VALOR DEPOSITO INICIAL
				            	getSetCEF.setValorOriginal(ret);
				            	
				            //	System.out.println("DEPOSITO INICIAL :>"+ret);
				          }

						 
						  
						  

//				          if(linhas.length-1 == i)//É O ULTIMO REGISTRO? SE SIM PEGA O VALOR 
//				          {
//				        	  	if(!linhas[i].equals("Voltar")) //Existem extratoq eu aultima linha tem o descricao voltar 
//				        	  		ret = linhas[i];
//				        	  	else
//				        	  		ret = linhas[i-1];
//				        	  	
//				        	  	ret =  trataSujeira(ret);
//				        	  	ret = ret.replace("C","");
//				        	  	ret = ret.replace("c","");
//				        	  	
//				        	  	//SETANDO O VALOR ATUALIZADO
//				        	  	getSetCEF.setValorAtualizado(ret.trim());
//				          }     
					  }
				  }	  
			  return ret;
		  }
		  
		  
		  public static String trataSujeira(String str) {
			  
			//  System.out.println("TRATAR SUJEIRA :|"+str+"|");
			  
		        String ret = "";
		        
		        ret = str;
		        
		        
		        ret = ret.replaceAll(" / ", "");// TRATA SUJEIRA DA CONTA JURIDICA
		        ret = ret.replaceAll("-", "");// TRATA SUJEIRA DA CONTA JURIDICA
		        ret = ret.replaceAll("CRED", "");//TRA SUJEIRA DO SALDO
		        ret = ret.replaceAll("DEP.DINH.", "");
		        ret = ret.replaceAll("Saldo Anterior", "");
		        
		        ret = ret.replaceAll("A VARA DO TRABALHO", "");
		        
           		ret = ret.replaceAll("[^\\p{ASCII}]", "");
           		ret = ret.replaceAll("//", "");
           		ret = ret.replaceAll("/", "");
           		ret = ret.replaceAll("#", "");
           		ret = ret.replaceAll("Saldodoperodo", "");
	        		ret = ret.replaceAll("Saldodoperodo", "");
	        		ret = ret.replaceAll("Anterio", "");

	        		ret = ret.replaceAll("..-", "");
        	  		ret = ret.replaceAll("Anterio", "");
        	  		ret = ret.replaceAll("DEP.DINH", "");
        	  		ret = ret.replaceAll("DEB.AUTOR", "");
        	  		ret = ret.replace("Remunerao", "");
        	  		
        	  		ret = ret.replace("Autor          :", "");
        	  		ret = ret.replace("         Saldo do período             ", "");
				ret = ret.replace("          Saldo do período            ", "");
				ret = ret.replace("    Saldo do periodo    ", "");
				ret = ret.replace("/", "");
				ret = ret.replace("                ", "");
				ret = ret.replace("Reclamante", "");
				ret = ret.replace(":", "");
				ret = ret.replace("     ", "");
				ret = ret.replace("     Saldo do periodo ", "");
				ret = ret.replace("Autor", "");
				ret = ret.replace("     Saldo do periodo    ", "");
				ret = ret.replace("     Saldo do periodo    ", "");
				ret = ret.replace("Saldo do periodo", "");
		       	ret = ret.replaceAll("[^\\p{ASCII}]", "");
	         	ret = ret.replaceAll("//", "");
	      		ret = ret.replaceAll("/", "");
	      		ret = ret.replaceAll("#", "");
	      		ret = ret.replaceAll("Saldodoperodo", "");
	     		ret = ret.replaceAll("Saldo do período", "");
	     		ret = ret.replaceAll("Saldo do perodo", "");
	     		ret = ret.replaceAll("      ", "");
		     		ret = ret.replace("CONTA JUDICIAL", "");
	       		return ret;
		    }
		  
		  
		  private static boolean campoNumerico(String campo){           
		        return campo.matches("[0-9]+");   
		}
				  
		  		  
		  public static String toTitledCase(String nome){
			  
			  //System.out.println("STR ENTRADA= "+ nome);
			  
			  nome = " "+nome; 
			  	
			  String aux =""; // só é utilizada para facilitar 

		        try{ //Bloco try-catch utilizado pois leitura de string gera a exceção abaixo
		            for(int i = 0; i < nome.length(); ++i){
		                if( nome.substring(i, i+1).equals(" ") || nome.substring(i, i+1).equals("  "))
		                {
		                    aux += nome.substring(i+1, i+2).toUpperCase();
		                   // System.out.println("1= "+ aux);
		                }
		                else
		                {
		                    aux += nome.substring(i+1, i+2).toLowerCase();
		                    //System.out.println("2= "+ aux);
		                }
		        }
		        }catch(IndexOutOfBoundsException indexOutOfBoundsException){
		            //não faça nada. só pare tudo e saia do bloco de instrução try-catch
		        }
		        nome = aux;
		       // System.err.println(nome);

			  return nome;
			}  
		  
		  
		  public static String diretorio(String caminho) {
			
			  File diretorio = new File(caminho); // ajfilho é uma pasta!
			  if (!diretorio.exists()) {
			     diretorio.mkdir(); //mkdir() cria somente um diretório, mkdirs() cria diretórios e subdiretórios.
			  } else {
			     System.out.println("Diretório já existente");
			  }

			  
			  return caminho;
			  
		  }
		  
		  public static String lerExcel(String str) throws IOException{

			  
			  
					int contAux = 0;// Controle para pegar o Numer da Conta Juridica
					int contadorPosicao = 11; //Para saber a linha que passou
					
					//Conta Juridica Existente:
					int posicaoExiste = 0; // Pega Guardar a posicao da Conta Existente
					int posicaoExisteUnica = 0; // Pega Guardar a posicao da Conta Existente
					
					boolean cjExiste = false;
					boolean cjExisteUnica = true;
					boolean cjNova = false;
					boolean barrarGravacaoContaJaExiste = false;
					
					long contaJudicial = 0;
					String valotAtualizado = "";
					String numeroParcela = "";
					boolean temValorParcela = false;
					String posicaoEvalor = "";
					
					String M = "M1";
					String L = "L1";
					String N = "N1";
					
					ArrayList<String> arrayConteudoExiste = new ArrayList<String> ();
					ArrayList<String> arrayConteudoNovoComValor = new ArrayList<String> ();
					ArrayList<String> arrayConteudoNovoSemValor = new ArrayList<String> ();
					ArrayList<String> arrayConteudoCJExisteNaoUnica = new ArrayList<String> ();
					
					ArrayList<String> arrayConteudoCJExisteNaoUnicaIncluirValor = new ArrayList<String> ();
					String ConteudoCJExisteNaoUnicaIncluir = "";
					
					ArrayList<String> arrayConteudoCJExisteNaoUnicaIncluirParcela = new ArrayList<String> ();
					String ConteudoExistenumeroParcela = "";
					String parcela = "";
					String CJXLX = "";
					String CJPDF = "";
			 
					try {

						ZipSecureFile.setMinInflateRatio(-1.0d);
			            XSSFWorkbook workbook = new XSSFWorkbook(excelBB);
			            XSSFSheet sheet = workbook.getSheetAt(0);
			            Row row = sheet.getRow(0);  
			            
			            getSetCEF.setCjExisteUnica(true);
			            getSetCEF.setCjExiste(false);
			            getSetCEF.setCjNova(true);
			            
			            CellReference cellReferencePAR = null;
			            Row rowLPAR = null;
			            Cell cellLPAR = null;
			            
			            CellReference cellReferenceVAL = null;
			            Row rowLVAL = null;
			            Cell cellLVAL = null;		           
			            
			            CellReference cellReferenceCJ = null;
			            Row rowLCJ = null;
			            Cell cellLCJ = null;
			            
			            
	
			            	for (int i = 11; i < 10000; i++) //Comeca do 11 para inicinar na linha 11
						{
			            	      try 
			            	      {
			            	    	  	//LENDO AS COLUNAS N1, N2 , N3  (VALOR ATUALIZADO)
//				            		 N = "N"+i;
//				            		 cellReferencePAR = new CellReference(N);   //Ferencia Coluna M usado na Conta Judicial
//				            		 rowLPAR = sheet.getRow(cellReferencePAR.getRow());	 //Ferencia Linha usado na Conta Judicial
//				            	     cellLPAR = rowLPAR.getCell(cellReferencePAR.getCol());
	
			            	    	  
			            	    	  
			            	    	  	//LENDO AS COLUNAS L1, L2 , L3  (VALOR ATUALIZADO)
				            		 L = "L"+i;
				            		 cellReferenceVAL = new CellReference(L);   //Ferencia Coluna M usado na Conta Judicial
				            		 rowLVAL = sheet.getRow(cellReferenceVAL.getRow());	 //Ferencia Linha usado na Conta Judicial
				            	     cellLVAL = rowLVAL.getCell(cellReferenceVAL.getCol());
	
			            	    	  
			            	    	  
			            	    	  	//LENDO AS COLUNAS M1, M2 , M3  (CONTA JUDICIAL)
				            		 M = "M"+i;			            		 
				            		 cellReferenceCJ = new CellReference(M);   //Ferencia Coluna M usado na Conta Judicial
				            		 rowLCJ = sheet.getRow(cellReferenceCJ.getRow());	 //Ferencia Linha usado na Conta Judicial
				            	     contaJudicial = 0;
		
			            	    	     cellLCJ = rowLCJ.getCell(cellReferenceCJ.getCol());  
				            	     
			            	    	     if(cellLCJ.CELL_TYPE_NUMERIC==cellLCJ.getCellType()) {
			                    		contaJudicial = (long) cellLCJ.getNumericCellValue();
				                    	contadorPosicao++;
				                    	getSetCEF.setContadorPosicao(contadorPosicao);
				            	     }
				            	     else if(cellLCJ.CELL_TYPE_STRING==cellLCJ.getCellType()) {
				            	    	 	CJXLX = cellLCJ.getStringCellValue();
				            	    	 	
				            	    	 	CJXLX = CJXLX.replace("//", "");
				            	    	 	CJXLX = CJXLX.replace("/", "");
				            	    	 	CJXLX = CJXLX.replace("-", "");
				            	    	 	CJXLX = CJXLX.replace("  ", "");
				            	    	 	CJXLX = CJXLX.trim();
				            	    	 	//System.out.println("CJ EXEL :>"+CJXLX);
				            	    	 	
				            	    	 	contaJudicial = Long.parseLong(CJXLX);
				                    	contadorPosicao++;
				                    	getSetCEF.setContadorPosicao(contadorPosicao);
				            	     }
			            	    	  
			            	    	     	CJPDF = getSetCEF.getContaJuridica();
				            	    	 	
			            	    	     	CJPDF = CJPDF.replace("//", "");
			            	    	     	CJPDF = CJPDF.replace("/", "");
			            	    	     	CJPDF = CJPDF.replace("-", "");
			            	    	     	CJPDF = CJPDF.replace("  ", "");
			            	    	     	CJPDF = CJPDF.trim();
				            	    	 	//System.out.println("CJ EXEL :>"+CJPDF);
			                    	//System.out.println("contadorPosicao: " + contadorPosicao);
				            	     

			            	    	     
			            	    	     
				            	      if(contaJudicial == Long.parseLong(CJPDF)) 
			                        	{
				            	    	  
				            	    	     System.out.println("contaJudicial EXEL :>"+contaJudicial);
				            	    	     System.out.println("contaJudicial PDF :>"+CJPDF);
				            	    	     System.out.println("POSICAO :>"+i);
				            	    	     
				            	    	     
				            	    	  		getSetCEF.setPosicaoExiste(i);	
			                        	 	getSetCEF.setCjNova(false);
	                  	 	
			                        	 
			                        	 	//---> VOU ADCIONANDO NO ARRAY ROSA QUE EXISTE PRA PINTAR DE UMA UNICA VEZ.
			                        	 	//arrayConteudoCJExisteNaoUnica.add(Integer.toString(getSetCEF.getPosicaoExiste()));
			                        	 	
			                        	 	
		                        	 		//Guarda o Valor atualizado para futura comparacao
			   			            	     if(cellLVAL.CELL_TYPE_NUMERIC==cellLVAL.getCellType()) {
			   			            	    	 	valotAtualizado = String.valueOf(cellLVAL.getNumericCellValue()) ;
						            	     }
			   			            	     if(cellLVAL.CELL_TYPE_STRING==cellLVAL.getCellType()) {
			   			            	    	 	valotAtualizado =	cellLVAL.getStringCellValue();
						            	     }
			   			            	     
//			   			            	     //Guardando a Parcela
//						            	     if(cellLPAR != null)
//						            	     {			            	    	 	
//						            	    	 	if(cellLPAR.CELL_TYPE_STRING==cellLPAR.getCellType() ) {			            	    	 		
//						            	    	 		numeroParcela = cellLPAR.getStringCellValue();
//						            	    	 	}
//						            	     }
			   			            	     
			   			            	     //---> ADIDIONA NESTE ARRAY PARA SABER SE JA EXISTE ANTES DE INCLUIR COM O MESMO VALOR E CONTA JUDICIAL
			   			            	  //	arrayConteudoCJExisteNaoUnicaIncluirValor.add(valotAtualizado);
			   			            	  //	arrayConteudoCJExisteNaoUnicaIncluirParcela.add(numeroParcela);
			   			            	     
			                        	 		
//			                        	 	if(getSetCEF.isCjExiste()) //ver se a conta existe e é unica ou nao
//			                        	 	{
//			                        	 		getSetCEF.setCjExisteUnica(false);
//			                        	 		//EscreverExistiNaoUnica();  // PINTA DE ROSA PARA AVIAR QUE TEM DUPLICIDADE E CONFERENCIA
//			                        	 		continue;// Se ja ter mais de uma para a verificacao		                        	 		
//			                        	 	}
		                        	 		
			                        	 	getSetCEF.setCjExiste(true); 
		                        	 		
			                        	}
	
							  }catch(Exception e){
								//  contadorPosicao ++;	
								  break;
						      }
						}
	
	
			            
			            if(getSetCEF.isCjNova()) 
		                {
				            	if(!getSetCEF.getValorAtualizado().equals("0,00"))// Se conter Valor
		                	 	{			                        	 		
		                	 		//System.out.println("3 - CONTA NAO EXISTE COM VALOR::" + getSetCEF.getValorAtualizado());
		                	 		EscreverContaNovaComValor();
		                	 	}
		                	 	else
		                	 	{
		                	 		//System.out.println("4 - CONTA NOVA VALOR ZERO:" + getSetCEF.getValorAtualizado());
		                	 		EscreverContaNovaValorZero();
		                	 	}
				            	
				            	//PAUSA PARA PODER GRAVAR O EXCEL
			  			     try {  
			  			        	//System.out.println(" PAUSA");
						        Thread.sleep( 100 );  
						     } 
			  			     catch (InterruptedException e) {  
						         e.printStackTrace();  
						     } 
		                }
			            
	      
			            
			            if(!getSetCEF.isCjNova() && getSetCEF.isCjExiste()) //Se existi é unica
	            	 		{
		            	 		//System.out.println("2-  CONTA EXISTE  É UNICA");
	
				            	if(!getSetCEF.getValorAtualizado().equals("0,00"))// Se conter Valor
		                	 	{			                        	 		
		                	 		//System.out.println("2.1 - VALOR ATUALIZADO: " + getSetCEF.getValorAtualizado());
		                	 		EscreverExistiUnicaComValor();   
		                	 	}
		                	 	else
		                	 	{
		                	 		//System.out.println("2.2 - VALOR ATUALIZADO ZERADO: " + getSetCEF.getValorAtualizado());
		                	 		EscreverExistiUnicaValorZero();  
		                	 	}
				            	
				            	
				            	//PAUSA PARA PODER GRAVAR O EXCEL
				            	try {  
						        Thread.sleep( 100 );  
						     } 
			  			     catch (InterruptedException e) {  
						         e.printStackTrace();  
						     } 
		            	 	}
//		            	 	else if(!getSetCEF.isCjExisteUnica()) 
//		            	 	{
//		            	 		String Valor = "";
//		            	 		String ValorApoio = "";
//		            	 		String valotVaiAtualizar = "";
//		            	 		
//		            	 		Valor  = getSetCEF.getValorAtualizado();
//		            	 		ValorApoio = getSetCEF.getValorAtualizado();
//		            	 		ValorApoio = ValorApoio.replace(".", "");
//		            	 		ValorApoio = ValorApoio.replace(",", "");
//	
//		            	 		
//		            	 		//COLOCO NO BEAN PARA QUANDO SABER AS LINHAS QUE PRECISA PINTAR DE ROSA QUE JA EXISTEM NAS LINHAS PASSADAS - EscreverExistiNaoUnica()
//		            	 		//getSetCEF.setArrayCJExisteNaoUnica(arrayConteudoCJExisteNaoUnica);
//		            	 		
//		            	 		//PINTA DE ROSA OS EXISTENTES  --- PINTAR ROS ESCURO sera os que nao conseguiu identificar
//		            	 		//EscreverExistiNaoUnica();
//	
//		            	 		//LIMPA O ARRAY PARA PODER PINTAR OS PROXIOS
//		            	 		//arrayConteudoCJExisteNaoUnica.clear();
//		            	 		
//		            	 		
//		            	 		
//		    			        for (int i = 0; i < arrayConteudoCJExisteNaoUnicaIncluirValor.size(); i++) {
//		    			        		
//		    			        		//Pegando Valor
//		    			        		valotVaiAtualizar = arrayConteudoCJExisteNaoUnicaIncluirValor.get(i);
//		    			        		valotVaiAtualizar = valotVaiAtualizar.replace(".", "");
//		    			        		valotVaiAtualizar = valotVaiAtualizar.replace(",", "");
//		    			        	
//		    			        		
//		    			        		//Pegando Parcela
//		    			        		ConteudoExistenumeroParcela = arrayConteudoCJExisteNaoUnicaIncluirParcela.get(i);
//		    			        		//ConteudoExistenumeroParcela = (ConteudoExistenumeroParcela.substring(ConteudoExistenumeroParcela.length() - 2).trim()).toUpperCase(); // Pega apenas os dois ultimos Digitos
//		    			        		
//		    			        		ConteudoExistenumeroParcela = (ConteudoExistenumeroParcela.toUpperCase()).trim();
//		    			        		ConteudoExistenumeroParcela = ConteudoExistenumeroParcela.replace("0", "");
//		    			        		
//		    			        		parcela = (getSetCEF.getParcela().toUpperCase()).trim();
//		    			        		parcela = parcela.replace("0", "");
//		    			        		
//		    			        		//Pegando a posicao e Valor para gravar caso a CJ existe e o numero da parcela tambem for igual
//		    			        		getSetCEF.setPosicaoExiste(Integer.parseInt(arrayConteudoCJExisteNaoUnica.get(i)));
//		    			        		
//		    			        		System.out.println("                                                           ");
//		    			        		System.out.println("Valor Excell         |" + valotVaiAtualizar+"|");	            	 		
//		    	            	 		System.out.println("Valor PDF            |" + ValorApoio +"|");
//		    	            	 		System.out.println("------------------------------------------------------------");
//		    	            	 		System.out.println("Parcela Excel        |" + ConteudoExistenumeroParcela +"|");	            	 		
//		    	            	 		System.out.println("Parcela PDF          |" + parcela +"|");	    			        		
//
//		    	            	 		System.out.println("                                                           ");
//			            	 		
//		    	            	 		
//		    	            	 		if(valotVaiAtualizar.equals(ValorApoio) || valotVaiAtualizar == ValorApoio) // Contas Juducuas ja sei que sao iguais a pergunra se o valor e diferente
//		    	                	 	{
//		    			            		System.out.println("VAI BARRAR -- VALORES IGUAIS NAO ATUALIZAR");
//		    			            		barrarGravacaoContaJaExiste = true;
//		    			            		break;
//		    	                	 	}
//			    	            	 	else if( ConteudoExistenumeroParcela.equals(parcela)   || ConteudoExistenumeroParcela == parcela) // Parcelas Iguais Atualiza Valor da Parcela
//			    	            	 	{
//			    	            	 		//SE ENCONTRAR A MESMA PARCELA
//			    	            	 		//PINTA DE ROSA E ATUALIZA O VALOR
//			    	            	 		
//			    	            	 		getSetCEF.setValorAtualizado(Valor);
//
//			    	            	 		System.out.println("CONTA IGUAL E PARCELA DIFERENTE-  VAI ATUALIZAR O VALOR");
//			    	            	 		EscreverExistiNaoUnicaComValor();
//			    	            	 		
//			    	            	 		barrarGravacaoContaJaExiste = true;
//			    	            	 		
//			    	            	 		break;
//			    	            	 	}
//			    	            	 	else
//			            	 				barrarGravacaoContaJaExiste = false;
//						        }
//			            	 		
//		
//			    			        if( !barrarGravacaoContaJaExiste)
//			    			        		EscreverExistiNaoUnicaNovo();
//	        	 				
//	        	 				
//			            	 	//PAUSA PARA PODER GRAVAR O EXCEL
//					        	try 
//					        	{  
//					        //System.out.println(" PAUSA");
//					        		Thread.sleep( 100 );  
//					        	} 
//					        	catch (InterruptedException e) {  
//							    e.printStackTrace();  
//							} 
//			            	 }	
		        } catch (IOException e) {
		            e.printStackTrace();
		        }				
				return "";
		  }
		  
		  
		  
		  public static void EscreverExistiUnicaComValor() throws IOException {
			  try{
				  	fo = new File(excelBB);
			        XSSFWorkbook a = new XSSFWorkbook(new FileInputStream(fo));
			        XSSFSheet my_sheet = a.getSheetAt(0);
			        
			        System.out.println("1-  EscreverExistiUnicaComValor GRAVAR NA LINHA :  " + getSetCEF.getPosicaoExiste());
			        

			        //Direita Cor Azul Claro
			        XSSFCellStyle style1 = a.createCellStyle();
			       // style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(0, 201, 222)));
			        style1.setAlignment ( XSSFCellStyle.ALIGN_RIGHT ) ; 
//			        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
//			        style1.setBorderBottom(CellStyle.BORDER_THIN);
//			        style1.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderLeft(CellStyle.BORDER_THIN);
//			        style1.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderRight(CellStyle.BORDER_THIN);
//			        style1.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderTop(CellStyle.BORDER_THIN);
//			        style1.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        //Centro Azul Claro
			        XSSFCellStyle style2 = a.createCellStyle();
			       // style2.setFillForegroundColor(new XSSFColor(new java.awt.Color(0, 201, 222)));
			        style2.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
//			        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
//			        style2.setBorderBottom(CellStyle.BORDER_THIN);
//			        style2.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderLeft(CellStyle.BORDER_THIN);
//			        style2.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderRight(CellStyle.BORDER_THIN);
//			        style2.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderTop(CellStyle.BORDER_THIN);
//			        style2.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        
			        CreationHelper createHelper = a.getCreationHelper();
			        XSSFCellStyle data = a.createCellStyle();
			        //data.setFillForegroundColor(new XSSFColor(new java.awt.Color(0, 201, 222)));
			        data.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
//			        data.setFillPattern(CellStyle.SOLID_FOREGROUND);
			        data.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yy"));
//			        data.setBorderLeft(CellStyle.BORDER_THIN);
//			        data.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        data.setBorderRight(CellStyle.BORDER_THIN);
//			        data.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        data.setBorderTop(CellStyle.BORDER_THIN);
//			        data.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        
			        //Centro Azul Claro
			        XSSFCellStyle style3 = a.createCellStyle();
			        style3.setFillForegroundColor(new XSSFColor(new java.awt.Color(0, 201, 222)));
			        style3.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
			        style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
			        style3.setBorderBottom(CellStyle.BORDER_THIN);
			        style3.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderLeft(CellStyle.BORDER_THIN);
			        style3.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderRight(CellStyle.BORDER_THIN);
			        style3.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderTop(CellStyle.BORDER_THIN);
			        style3.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			       


//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(1).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(2).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(3).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(4).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(5).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(6).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(7).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(8).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(9).setCellStyle(data);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(10).setCellStyle(style1);
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(11).setCellValue(getSetCEF.getValorAtualizado());			        
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(11).setCellStyle(style1);
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);
			        
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(12).setCellStyle(style3);
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(12).setCellType(XSSFCell.CELL_TYPE_STRING);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(13).setCellStyle(style2);
			        
			        FileOutputStream outputStream = new FileOutputStream(new File(excelBB));
			        a.write(outputStream);
			        outputStream.close();//Close in finally if possible
			        outputStream = null;


  
			        
		        }catch(Exception e){
		        }
			}

		  
		  public static void EscreverExistiUnicaValorZero() throws IOException {
			  try{
				  	fo = new File(excelBB);
			        XSSFWorkbook a = new XSSFWorkbook(new FileInputStream(fo));
			        XSSFSheet my_sheet = a.getSheetAt(0);
			        
			        System.out.println("2 -  EscreverExistiUnicaValorZero GRAVAR NA LINHA :  " + getSetCEF.getPosicaoExiste());
			        

			        //Direita Cor Azul Claro
			        XSSFCellStyle style1 = a.createCellStyle();
			        //style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(0, 251, 254)));
			        style1.setAlignment ( XSSFCellStyle.ALIGN_RIGHT ) ; 
//			        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
//			        style1.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderLeft(CellStyle.BORDER_THIN);
//			        style1.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderRight(CellStyle.BORDER_THIN);
//			        style1.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderTop(CellStyle.BORDER_THIN);
//			        style1.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        //System.out.println("Passo1");
			        
			        //Centro Azul Claro
			        XSSFCellStyle style2 = a.createCellStyle();
			       // style2.setFillForegroundColor(new XSSFColor(new java.awt.Color(0, 251, 254)));
			        style2.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
//			        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
//			        style2.setBorderBottom(CellStyle.BORDER_THIN);
//			        style2.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderLeft(CellStyle.BORDER_THIN);
//			        style2.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderRight(CellStyle.BORDER_THIN);
//			        style2.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderTop(CellStyle.BORDER_THIN);
//			        style2.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        //System.out.println("Passo2");
			        
			        
			        CreationHelper createHelper = a.getCreationHelper();
			        XSSFCellStyle data = a.createCellStyle();
			        //data.setFillForegroundColor(new XSSFColor(new java.awt.Color(0, 251, 254)));
			        data.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
//			        data.setFillPattern(CellStyle.SOLID_FOREGROUND);
			        data.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yy"));
//			        data.setBorderLeft(CellStyle.BORDER_THIN);
//			        data.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        data.setBorderRight(CellStyle.BORDER_THIN);
//			        data.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        data.setBorderTop(CellStyle.BORDER_THIN);
//			        data.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        
			        //Centro Azul Claro
			        XSSFCellStyle style3 = a.createCellStyle();
			        style3.setFillForegroundColor(new XSSFColor(new java.awt.Color(183, 16, 46)));
			        style3.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
			        style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
			        style3.setBorderBottom(CellStyle.BORDER_THIN);
			        style3.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderLeft(CellStyle.BORDER_THIN);
			        style3.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderRight(CellStyle.BORDER_THIN);
			        style3.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderTop(CellStyle.BORDER_THIN);
			        style3.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));

			        
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(1).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(2).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(3).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(4).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(5).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(6).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(7).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(8).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(9).setCellStyle(data);			        
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(10).setCellStyle(style1);
			        
			        //Nao atualziar valor quando for zerado ja existe
			        //my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(11).setCellValue(getSetCEF.getValorAtualizado());			        
			        //my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(11).setCellStyle(style1);
			        //my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);
			        
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(12).setCellStyle(style3);
			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(12).setCellType(XSSFCell.CELL_TYPE_STRING);
//			        my_sheet.getRow(getSetCEF.getPosicaoExiste()-1).getCell(13).setCellStyle(style2);
			        
     
			        FileOutputStream outputStream = new FileOutputStream(new File(excelBB));
			        a.write(outputStream);
			        outputStream.close();//Close in finally if possible
			        
			        outputStream = null;
  
			        
		        }catch(Exception e){
		        }
			}
		  
		  
		  public static void EscreverContaNovaComValor() throws IOException {
			  try{
				  
				  
				  	
			        XSSFWorkbook a = new XSSFWorkbook(new FileInputStream(fo));
			        
			        XSSFSheet my_sheet = a.getSheetAt(0);
			        
			        
			        System.out.println("3 -  EscreverContaNovaComValor GRAVAR NA LINHA :  " + getSetCEF.getContadorPosicao());
			        

			        //Direita Cor Azul Claro
			        XSSFCellStyle style1 = a.createCellStyle();
			       // style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(89, 170, 8)));
			        style1.setAlignment ( XSSFCellStyle.ALIGN_RIGHT ) ; 
//			        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
//			        style1.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderLeft(CellStyle.BORDER_THIN);
//			        style1.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderRight(CellStyle.BORDER_THIN);
//			        style1.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderTop(CellStyle.BORDER_THIN);
//			        style1.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        
			        //Centro Azul Claro
			        XSSFCellStyle style2 = a.createCellStyle();
			    //    style2.setFillForegroundColor(new XSSFColor(new java.awt.Color(89, 170, 8)));
			        style2.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
//			        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
//			        style2.setBorderBottom(CellStyle.BORDER_THIN);
//			        style2.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderLeft(CellStyle.BORDER_THIN);
//			        style2.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderRight(CellStyle.BORDER_THIN);
//			        style2.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderTop(CellStyle.BORDER_THIN);
//			        style2.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        
			        
			        //Centro Com data Azul
			        CreationHelper createHelper = a.getCreationHelper();
			        XSSFCellStyle data = a.createCellStyle();
			        //data.setFillForegroundColor(new XSSFColor(new java.awt.Color(89, 179, 8)));
			        data.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
//			        data.setFillPattern(CellStyle.SOLID_FOREGROUND);
			        data.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yy"));
//			        data.setBorderLeft(CellStyle.BORDER_THIN);
//			        data.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        data.setBorderRight(CellStyle.BORDER_THIN);
//			        data.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        data.setBorderTop(CellStyle.BORDER_THIN);
//			        data.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        
			        //Centro Azul Claro
			        XSSFCellStyle style3 = a.createCellStyle();
			        style3.setFillForegroundColor(new XSSFColor(new java.awt.Color(89, 179, 8)));
			        style3.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
			        style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
			        style3.setBorderBottom(CellStyle.BORDER_THIN);
			        style3.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderLeft(CellStyle.BORDER_THIN);
			        style3.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderRight(CellStyle.BORDER_THIN);
			        style3.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderTop(CellStyle.BORDER_THIN);
			        style3.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        


			        my_sheet.createRow(getSetCEF.getContadorPosicao()-1);

			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(1);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(1).setCellValue(getSetCEF.getAutor());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(1).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(2);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(2).setCellValue(getSetCEF.getReu());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(2).setCellStyle(style2);
		        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(3);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(3).setCellValue(getSetCEF.getCNPJ());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(3).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(4);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(4).setCellValue(getSetCEF.getProcesso());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(4).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(5);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(5).setCellValue(getSetCEF.getVara());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(5).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(6);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(6).setCellValue(getSetCEF.getComarca());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(6).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(7);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(7).setCellValue(getSetCEF.getEstado());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(7).setCellStyle(style2);

			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(8);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(8).setCellValue("Trabalhista");
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(8).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(9);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(9).setCellValue(getSetCEF.getDataDeposito());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(9).setCellStyle(data);

			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(10);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(10).setCellValue(getSetCEF.getValorOriginal());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(10).setCellStyle(style1);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(10).setCellType(XSSFCell.CELL_TYPE_STRING);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(11);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(11).setCellValue(getSetCEF.getValorAtualizado());			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(11).setCellStyle(style1);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(12);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(12).setCellValue(getSetCEF.getContaJuridica());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(12).setCellStyle(style3);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(12).setCellType(XSSFCell.CELL_TYPE_STRING);
			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(13);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(13).setCellValue(getSetCEF.getParcela());			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(13).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(13).setCellType(XSSFCell.CELL_TYPE_STRING);

			        FileOutputStream outputStream = new FileOutputStream(new File(excelBB));
			        a.write(outputStream);
			        outputStream.close();//Close in finally if possible
			        outputStream = null;
			        
		        }catch(Exception e){
		        		System.out.println("3 -  EscreverContaNovaComValor GRAVAR NA LINHA EROOO");
		        		System.out.println(" ERRO: " + e);
		        }
			}	
		  		 
		  
		  public static void EscreverContaNovaValorZero() throws IOException {
			  try{
				  	fo = new File(excelBB);
			        XSSFWorkbook a = new XSSFWorkbook(new FileInputStream(fo));
			        
			        XSSFSheet my_sheet = a.getSheetAt(0);
			        
			        
			        System.out.println("4 -  EscreverContaNovaValorZero GRAVAR NA LINHA :  " + getSetCEF.getContadorPosicao());
			        

			        //Direita Cor Azul Claro
			        XSSFCellStyle style1 = a.createCellStyle();
			        //style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(9, 232, 5)));
			        style1.setAlignment ( XSSFCellStyle.ALIGN_RIGHT ) ; 
//			        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
//			        style1.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderLeft(CellStyle.BORDER_THIN);
//			        style1.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderRight(CellStyle.BORDER_THIN);
//			        style1.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style1.setBorderTop(CellStyle.BORDER_THIN);
//			        style1.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        
			        //Centro Azul Claro
			        XSSFCellStyle style2 = a.createCellStyle();
			       // style2.setFillForegroundColor(new XSSFColor(new java.awt.Color(9, 232, 5)));
			        style2.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
//			        style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
//			        style2.setBorderBottom(CellStyle.BORDER_THIN);
//			        style2.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderLeft(CellStyle.BORDER_THIN);
//			        style2.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderRight(CellStyle.BORDER_THIN);
//			        style2.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        style2.setBorderTop(CellStyle.BORDER_THIN);
//			        style2.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        
			        
			        //Centro Com data Azul
			        CreationHelper createHelper = a.getCreationHelper();
			        XSSFCellStyle data = a.createCellStyle();
			        //data.setFillForegroundColor(new XSSFColor(new java.awt.Color(9, 232, 5)));
			        data.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
//			        data.setFillPattern(CellStyle.SOLID_FOREGROUND);
			        data.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yy"));
//			        data.setBorderLeft(CellStyle.BORDER_THIN);
//			        data.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        data.setBorderRight(CellStyle.BORDER_THIN);
//			        data.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
//			        data.setBorderTop(CellStyle.BORDER_THIN);
//			        data.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			        
			        //Centro Azul Claro
			        XSSFCellStyle style3 = a.createCellStyle();
			        style3.setFillForegroundColor(new XSSFColor(new java.awt.Color(9, 232, 5)));
			        style3.setAlignment ( XSSFCellStyle.ALIGN_CENTER ) ; 
			        style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
			        style3.setBorderBottom(CellStyle.BORDER_THIN);
			        style3.setBottomBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderLeft(CellStyle.BORDER_THIN);
			        style3.setLeftBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderRight(CellStyle.BORDER_THIN);
			        style3.setRightBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        style3.setBorderTop(CellStyle.BORDER_THIN);
			        style3.setTopBorderColor(new XSSFColor(new java.awt.Color(0, 0, 0)));
			        
			       
			        my_sheet.createRow(getSetCEF.getContadorPosicao()-1);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(1);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(1).setCellValue(getSetCEF.getAutor());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(1).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(2);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(2).setCellValue(getSetCEF.getReu());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(2).setCellStyle(style2);
		        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(3);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(3).setCellValue(getSetCEF.getCNPJ());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(3).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(4);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(4).setCellValue(getSetCEF.getProcesso());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(4).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(5);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(5).setCellValue(getSetCEF.getVara());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(5).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(6);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(6).setCellValue(getSetCEF.getComarca());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(6).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(7);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(7).setCellValue(getSetCEF.getEstado());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(7).setCellStyle(style2);

			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(8);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(8).setCellValue("Trabalhista");
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(8).setCellStyle(style2);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(9);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(9).setCellValue(getSetCEF.getDataDeposito());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(9).setCellStyle(data);

			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(10);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(10).setCellValue(getSetCEF.getValorOriginal());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(10).setCellStyle(style1);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(10).setCellType(XSSFCell.CELL_TYPE_STRING);
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(11);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(11).setCellValue(getSetCEF.getValorAtualizado());			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(11).setCellStyle(style1);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(11).setCellType(XSSFCell.CELL_TYPE_STRING);
			        
			        
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(12);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(12).setCellValue(getSetCEF.getContaJuridica());
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(12).setCellStyle(style3);
			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(12).setCellType(XSSFCell.CELL_TYPE_STRING);
			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).createCell(13);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(13).setCellValue(getSetCEF.getParcela());			        
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(13).setCellStyle(style2);
//			        my_sheet.getRow(getSetCEF.getContadorPosicao()-1).getCell(13).setCellType(XSSFCell.CELL_TYPE_STRING);

			        
			        FileOutputStream outputStream = new FileOutputStream(new File(excelBB));
			        a.write(outputStream);
			        outputStream.close();//Close in finally if possible
			        
			        outputStream = null;
  
			        
		        }catch(Exception e){
		        		System.out.println("4-  EscreverContaNovaComValor GRAVAR NA LINHA EROOO");
		        		System.out.println(" ERRO: " + e);
		        }
			}

	  
		  
}