package com.macroproyectos.plantillas;

import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.KeyStore;
import java.security.Security;
import java.security.cert.X509Certificate;
import java.util.ArrayList;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.Vector;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import javax.imageio.ImageIO;
import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.ConsoleAppender;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.log4j.PatternLayout;
import org.bouncycastle.jce.provider.BouncyCastleProvider;
import org.w3c.dom.NodeList;

import com.aspose.words.Bookmark;
import com.aspose.words.BookmarkCollection;
import com.aspose.words.BookmarkEnd;
import com.aspose.words.BookmarkStart;
import com.aspose.words.Cell;
import com.aspose.words.CellCollection;
import com.aspose.words.CertificateHolder;
import com.aspose.words.CompositeNode;
import com.aspose.words.CustomDocumentProperties;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.DocumentProperty;
import com.aspose.words.FontSettings;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.License;
import com.aspose.words.LoadOptions;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeImporter;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.ProtectionType;
import com.aspose.words.Range;
import com.aspose.words.Row;
import com.aspose.words.SaveFormat;
import com.aspose.words.Section;
import com.aspose.words.SectionCollection;
import com.aspose.words.Shape;
import com.aspose.words.Table;
import com.macroproyectos.plantillas.errores.PlantillasAPIException;
import com.macroproyectos.plantillas.logger.PlantillasLogger;

import net.sf.saxon.dom.DOMNodeList;
import net.sf.saxon.lib.NamespaceConstant;

/**
 * Clase para el manejo de documentos de Word
 *
 * @author Carlos Mario
 */
public class DocxTemplate implements DocumentTemplates {

    private static final Logger logger = Logger.getLogger(DocxTemplate.class);
    private boolean isOpen = false;
    private com.aspose.words.Document doc;
    private ArrayList<BookMarkProperties> bkProps;
    private HashMap<String, String> invalidTagsBk = new HashMap<>();

    /**
     * Carga un archivo de Word para manipular con esta API.
     * 
     * @param passwordDoc contrasena del documento en caso de que tenga
     * @param file Un arreglo de bytes que representa el archivo a cargar
     * 
     */
    @Override
    public void load(byte[] file) {
        open(file);
    }

    /**
     * Guarda los cambios ejecutados en el archivo y devuelve un arreglo de
     * bytes con el archivo ya modificado.
     *
     * @return Un arreglo de bytes con el archivo ya modificado
     * @throws Exception
     */
    @Override
    public byte[] getFile() {
        doc.removeMacros();

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            doc.save(out, SaveFormat.DOCX);
            return out.toByteArray();
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    /**
     * Abre un archivo dada una ubicacion en disco
     *
     * @param path La ruta donde se haya el archivo de Word
     * 
     */
    @Override
    public void open(String path) {
        try {
            open(FileUtils.readFileToByteArray(new java.io.File(path)));
        } catch (IOException e) {
            logger.error("Error IOException", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    /**
     * Abre un archivo dado un arreglo de bytes que representa su contenido
     *
     * @param bytes El arreglo de bytes que contiene el archivo
     * 
     */
    public void open(byte[] bytes) {
        try {
            License license = new License();
            InputStream lStream = this.getClass().getResourceAsStream("/Aspose.Total.Java.lic");
            logger.debug(lStream.toString());
            license.setLicense(lStream);
            doc = new com.aspose.words.Document(new ByteArrayInputStream(bytes));
            obtenerMarcasDoc();

            // Cargar las marcas invÃ¡lidas
            invalidTagsBk.put("tbl", "contiene una tabla");
            invalidTagsBk.put("bookmarkStart", "contiene un inicio de marca");
            invalidTagsBk.put("bookmarkEnd", "contiene un fin de marca");

            isOpen = true;

        } catch (Exception ex) {
            logger.error("Error general", ex);
            if (ex.getMessage().length() == 0 && ex.getMessage().contains("from ZipFile")) {
                throw new PlantillasAPIException(1);
            }

            if (ex.getMessage().contains("Unknown file format")) {
                throw new PlantillasAPIException(1);
            }

            throw new PlantillasAPIException(10, ex);
        }

    }

    /**
     * Recorre el documento y localiza todas las marcas que se hayan definido.
     * Excluira las marcas de tabla de contenido, vinculos OLE y la etiqueta
     * especial _GoBack.
     *
     * @throws Exception
     */
    private void obtenerMarcasDoc() throws Exception {
        try {
            BookmarkCollection bookmarks = doc.getRange().getBookmarks();
            bkProps = new ArrayList<>();

            for (Bookmark bk : bookmarks) {

                if ("_GoBack".equals(bk.getName()) || bk.getName().startsWith("_Toc")
                        || bk.getName().startsWith("OLE_")) {

                    try {
                        if ("_GoBack".equals(bk.getName())) {
                            bk.getBookmarkEnd().remove();
                        }
                    } catch (Exception e) {
                        logger.info("getBookmarkEnd no encontrado");
                    }

                    continue;
                }

                BookMarkProperties curr = new BookMarkProperties();
                curr.setName(bk.getName());
                curr.setType(getTypeBookMark(curr, bk.getName()));

                Node ancestor = bk.getBookmarkStart().getAncestor(NodeType.TABLE);
                if (ancestor != null) {
                    curr.setIdTable("" + ancestor.hashCode());
                }

                bkProps.add(curr);
            }
        } catch (Exception ex) {
            logger.error("Error general", ex);
            throw new PlantillasAPIException(10, ex);
        }

    }

    /**
     * Obtiene de que tipo es una marca en el documento (text o img)
     *
     * @param bkName El nombre de la marca
     * @return Una cadena con el tipo de marca (text o img)
     * 
     */
    private String getTypeBookMark(BookMarkProperties prop, String bkName) {
        Bookmark bk = doc.getRange().getBookmarks().get(bkName);

        Node node = bk.getBookmarkStart();
        int level = 0;

        while ((node = node.getNextSibling()) != null) {
            if (++level > 10) {
                // controlar maximo iteraciones
                break;
            }
            if (node instanceof BookmarkEnd) {

                // Si es una marca de fin se verifica que sea la de la marca actual
                BookmarkEnd bookmarkEnd = (BookmarkEnd) node;
                try {
                    if (bookmarkEnd.equals(bk.getBookmarkEnd())) {
                        // si la marca de fin coincide con la inicial fin del ciclo de verificacion ...
                        break;
                    }
                } catch (Exception e) {
                    logger.info("Excepcion leyendo marca final", e);
                }

            } else if (node instanceof Shape) {
                Shape currentShape = (Shape) node;
                double width = currentShape.getWidth();
                double height = currentShape.getHeight();
                prop.setWidth(width);
                prop.setHeight(height);
                return "img";
            }
        }

        logger.info("[" + bkName + "] es tipo texto " + bk.getBookmarkStart().getNextSibling());
        return "text";

    }

    /**
     * Obtiene un arreglo de BookMarkProperties con todas las marcas que se
     * encuentren en el documento.
     *
     * @return null si el documento no estï¿½ abierto, de lo contrario, un
     *         ArrayList de todas las marcas
     */
    @Override
    public List<BookMarkProperties> getMarksNames() {
        if (!isOpen) {
            return null;
        }
        return bkProps;
    }

    public BookMarkProperties findBookmark(String name) {
        for (BookMarkProperties prop : getMarksNames()) {
            if (prop.getName().equalsIgnoreCase(name)) {
                return prop;
            } else {
                logger.info(
                        "__________________La marca ,{}'" + prop.getName() + "' no esta en el documento , {}" + name);
            }
        }
        return null;
    }

    /**
     * Obtiene el texto actual de una marca en el documento
     *
     * @param marca El nombre de la marca
     * @return El valor actual contenido de la marca
     * @throws Exception
     */
    public String getMarkText(String marca) throws Exception {
        Bookmark bk = doc.getRange().getBookmarks().get(marca);

        if (bk == null) {
            logger.info("La marca '" + marca + "' no esta en el documento");
            return null;
        }

        return bk.getText();
    }

    /**
     * Asigna un texto a una marca en el documento
     *
     * @param marca El nombre de la marca
     * @param valor El contenido que se desea incluir en la marca
     * @throws Exception
     */
    @Override
    public void combineMarkText(String marca, String valor) {
    	removeDuplicatedMarks(marca);
        Bookmark bk = doc.getRange().getBookmarks().get(marca);

        if (bk == null) {
            logger.info("La marca '" + marca + "' no esta en el documento");
            return;
        }

        logger.info("Marca :'" + marca + "' = " + valor);
        try {
            bk.setText(valor);
        } catch (IllegalArgumentException exc) {
            PlantillasLogger.mainLogger.error("Error en marca " + marca, exc);
            throw new PlantillasAPIException(8, marca);
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    /**
     * Asigna una imagen a una marca en el documento
     *
     * @param marca El nombre de la marca
     * @param valor Un arreglo de bytes que representa la imagen que se desea
     *            incluir
     * @param awidth ancho de la imagen a combinar           
     * @param aheight altura de la imagen a combinar
     */
    public void combineMarkImg(String marca, byte[] valor, int awidth, int aheight) {
    	removeDuplicatedMarks(marca);
        logger.info("Marca de imagen actual: " + marca);
        Bookmark bk = null;
        bk = doc.getRange().getBookmarks().get(marca);

        if (bk == null) {
            logger.info("La marca '" + marca + "' no esta en el documento");
            return;
        }

        BookMarkProperties prop = findBookmark(marca);

        try {
            bk.setText("");
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.moveTo(bk.getBookmarkStart());
            bk.remove();

            builder.startBookmark(marca);
            double width = awidth>0? awidth: prop.getWidth();
            double height = aheight>0? aheight : prop.getHeight();

            // Las marcas no se detectaron correctamente => identificar el ancho y el alto a partir de la imagen
            if (width == 0 || height == 0) {
                InputStream in = new ByteArrayInputStream(valor);

                BufferedImage buf = ImageIO.read(in);

                height = buf.getHeight();
                width = buf.getWidth();

                in.close();
            }

            builder.insertImage(valor, width, height);
            builder.endBookmark(marca);
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }
    /**
     * Asigna una imagen a una marca en el documento
     *
     * @param marca El nombre de la marca
     * @param valor Un arreglo de bytes que representa la imagen que se desea
     *            incluir
     * 
     */
    @Override
    public void combineMarkImg(String marca, byte[] valor) {
     combineMarkImg(marca, valor, 0, 0);    
    }

    
    /**
     * Asigna contenido HTML a una marca en el documento
     *
     * @param marca El nombre de la marca
     * @param valor El contenido del documento HTML
     * @throws Exception
     */
    @Override
    public void combineMarkHTML(String marca, String valor) {
		removeDuplicatedMarks(marca);
        if (PlantillasLogger.mainLogger.isDebugEnabled()) {
            String tempPath = System.getProperty("java.io.tmpdir") + "/";
            try {
                FileUtils.writeByteArrayToFile(new File(tempPath + System.currentTimeMillis() + ".html"),
                        valor.getBytes());
            } catch (IOException e) {
                logger.error("Error IOException", e);
                throw new PlantillasAPIException(10, e);
            }
        }

        Bookmark bk = doc.getRange().getBookmarks().get(marca);

        if (bk == null) {
            logger.info("La marca '" + marca + "' no esta en el documento");
            return;
        }
           
        try {
            bk.setText("");
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.moveTo(bk.getBookmarkStart());
            bk.remove();

            builder.startBookmark(marca);
            builder.writeln();
            builder.insertHtml(valor);
            builder.writeln();
            builder.endBookmark(marca);
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    /**
     * remueve las marcas duplicadas en el documento (inconsistencias de libre office)
     * Se detecto que entre la primera marca y la ultima marca (BookmarkStart,BookmarkEnd) quedaban etiquetas de marcas  sueltas o repetidas.
     * El algoritmo toma la primera marca BookmarkStart y la ultima BookmarkEnd y borra todos los elementos entre ellas , antes de hacer la combinación   
     * Esta operación solo se realiza cuando  se detecta una marca repetida
     * @param marca nombre de marca a validar
     * @return
     */
	private void removeDuplicatedMarks(String marca) {
		BookmarkCollection bml = doc.getRange().getBookmarks();
        List <Bookmark>referenciaMarca = new ArrayList<>();
        for(Bookmark bm1:bml) {
        	if(bm1.getName().equals(marca))
        	  {
        		referenciaMarca.add(bm1);
        	  }        	
        } 
        
        DocumentBuilder builder = new DocumentBuilder(doc);
        if(referenciaMarca.size()>1) {
          
        	   Node next = referenciaMarca.get(0).getBookmarkStart().getNextSibling();
        	   Node node = referenciaMarca.get(0).getBookmarkStart();
        	   while (next!= null) {
        		   if(next instanceof BookmarkStart) {
        			   BookmarkStart bms= (BookmarkStart) next;
        			   if(bms.getName().equals(marca)){
        				   referenciaMarca.remove(0);
        				   node.remove();
        				   node = next;
        				   next = next.getNextSibling();
        				   continue;
        			   }
        			   else {
        				   logger.error("Inconsistencia de documento, marca '"+marca+"' sin Final  ...  ");
        				   break;
        			   }
        		   }
        		   else if(next instanceof BookmarkEnd && referenciaMarca.size()==1) {
        			   logger.info("ID END: " + ((BookmarkEnd)next).toString());
        			   break;
        			 
        		   }
        		   try {
        		   Node old = next;
        		   next = next.getNextSibling();
        		   old.remove();
        		   }catch(Exception t) {
        			   logger.warn("Error removiendo contenido api Aspose: " ,t);
        		   }
        	   }
        }
	}

    /**
     * Basado en un conjunto de encabezados y datos, construye dentro de una
     * marca una tabla en Word con toda la informaciï¿½n suministrada.
     *
     * @param marca El nombre de la marca
     * @param headers Un arreglo con los encabezados de la tabla
     * @param datagrid Una matriz de datos, que contiene la informacion a
     *            desplegar en la tabla
     * 
     */
    @Override
    public void combineMarkGrid(String marca, List<String> headers, List<ArrayList> datagrid) {
    	removeDuplicatedMarks(marca);
        Bookmark bookmark = doc.getRange().getBookmarks().get(marca);

        try {
            bookmark.setText("");
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.moveTo(bookmark.getBookmarkStart());
            bookmark.remove();

            builder.startBookmark(marca);
            builder.startTable();

            // Pintar encabezados
            Color previousFontColor = builder.getFont().getColor();
            Color previousBgColor = builder.getCellFormat().getShading().getBackgroundPatternColor();
            builder.getCellFormat().getShading().setBackgroundPatternColor(Color.blue);
            builder.getFont().setColor(Color.white);

            for (String header : headers) {
                builder.insertCell();
                builder.write(header);
            }

            builder.endRow();

            builder.getCellFormat().getShading().setBackgroundPatternColor(previousBgColor);
            builder.getFont().setColor(previousFontColor);

            for (ArrayList row : datagrid) {
                logger.info(row);
                for (Object cell : row) {

                    builder.insertCell();
                    if (cell == null) {
                        continue;
                    }
                    builder.write(cell.toString());
                }
                builder.endRow();
            }

            builder.endTable();
            builder.endBookmark(marca);
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    private void insertDocument(Node insertAfterNode, Document srcDoc, String marca) throws Exception {
    	removeDuplicatedMarks(marca);
        // Make sure that the node is either a paragraph or table.
        if ((insertAfterNode.getNodeType() != NodeType.PARAGRAPH)
                && (insertAfterNode.getNodeType() != NodeType.TABLE)) {
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }

        Bookmark bookmark = doc.getRange().getBookmarks().get(marca);

        try {
            bookmark.setText("");
        } catch (Exception exc) {
            logger.error("Error general", exc);
        }

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveTo(bookmark.getBookmarkStart());
        bookmark.remove();

        // We will be inserting into the parent of the destination paragraph.
        CompositeNode dstStory = insertAfterNode.getParentNode();

        // This object will be translating styles and lists during the import.
        NodeImporter importer = new NodeImporter(srcDoc, insertAfterNode.getDocument(),
                ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Loop through all sections in the source document.
        builder.startBookmark(marca);
        Node lastNode = null;
        for (Section srcSection : srcDoc.getSections()) {
            // Loop through all block level nodes (paragraphs and tables) in the body of the section.
            for (Node srcNode : (Iterable<Node>) srcSection.getBody()) {
                // Let's skip the node if it is a last empty paragraph in a section.
                if (srcNode.getNodeType() == (NodeType.PARAGRAPH)) {
                    Paragraph para = (Paragraph) srcNode;
                    if (para.isEndOfSection() && !para.hasChildNodes()) {
                        continue;
                    }
                }

                // This creates a clone of the node, suitable for insertion into the destination document.
                Node newNode = importer.importNode(srcNode, true);

                // Insert new node after the reference node.
                lastNode = dstStory.insertAfter(newNode, insertAfterNode);
                insertAfterNode = newNode;
            }
        }

        lastNode = dstStory.insertAfter(new Paragraph(doc), insertAfterNode);

        logger.debug(lastNode);
        builder.moveTo(lastNode);
        builder.endBookmark(marca);
    }

    /**
     * Asigna otro documento de Word al contenido de una marca
     *
     * @param marca El nombre de la marca
     * @param valor El contenido del documento a incluir
     * 
     */
    @Override
    public void combineMarkDoc(String marca, byte[] valor) {
    	removeDuplicatedMarks(marca);
        if (PlantillasLogger.mainLogger.isDebugEnabled()) {
            String tempPath = System.getProperty("java.io.tmpdir") + "/";
            try {
                FileUtils.writeByteArrayToFile(new File(tempPath + System.currentTimeMillis() + ".docx"), valor);
            } catch (IOException e) {
                logger.error("Error IOException", e);
                throw new PlantillasAPIException(10, e);
            }
        }

        try {
            Document subDoc = new Document(new ByteArrayInputStream(valor));

            Bookmark bookmark = doc.getRange().getBookmarks().get(marca);

            if (bookmark == null) {
                logger.info("La marca '" + marca + "' no esta en el documento");
                return;
            }

            CompositeNode parent = bookmark.getBookmarkStart().getParentNode();

            insertDocument(parent, subDoc, marca);
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    /**
     * Habilita el control de cambios dentro del documento
     */
    @Override
    public void enableTrackChanges() {
        doc.setTrackRevisions(true);
    }

    /**
     * Borra todos los comentarios dentro de un documento
     * 
     */
    @Override
    public void deleteAllComments() {
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        comments.clear();
    }

    /**
     * Acepta todos las revisiones del documento
     *
     * 
     */
    @Override
    public void acceptAllChanges() {
        try {
            doc.acceptAllRevisions();
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    /**
     * Deshabilita el control de cambios en el documento
     */
    @Override
    public void disableTrackChanges() {
        doc.setTrackRevisions(false);
    }

    /**
     * Toma todas las secciones del documento y las protege con el metodo
     * "Rellenando Formularios"
     */
    @Override
    public void lockAllSections() {
        SectionCollection sections = doc.getSections();

        for (Section section : sections) {
            section.setProtectedForForms(true);
        }
    }

    private boolean hasVersion = false;

    /**
     * Asigna al documento una versiï¿½n para propï¿½sitos de verificaciï¿½n interna
     *
     * @param ver El valor que se desea usar como versiï¿½n
     */
    @Override
    public void setVersion(String ver) {
        CustomDocumentProperties custom = doc.getCustomDocumentProperties();

        if (!hasVersion) {
            custom.add("_Version", ver);
            hasVersion = true;
        }

    }

    /**
     * Recupera el valor actual de la version del documento
     *
     * @return El valor actual de la version
     */
    @Override
    public String getVersion() {
        CustomDocumentProperties custom = doc.getCustomDocumentProperties();

        for (DocumentProperty property : custom) {
            logger.info(property.getName());
        }

        if (custom.get("_Version") != null) {
            return custom.get("_Version").getValue().toString();
        }

        return null;
    }

    /**
     * Crea una propiedad personalizada dentro del documento, si esta propiedad
     * ya existe, reemplaza su valor con el suministrado.
     *
     * @param name El nombre de la propiedad
     * @param val El valor de la propiedad
     */
    @Override
    public void setProperty(String name, String val) {
        CustomDocumentProperties custom = doc.getCustomDocumentProperties();
        if (custom.get(name) == null) {
            if (val != null) {
                custom.add(name, val);
            }
        } else if (val == null) {
            custom.remove(name);
        } else {
            custom.get(name).setValue(val);
        }
    }

    /**
     * Recupera el valor de una propiedad
     *
     * @param name El nombre de la propiedad
     * @return El valor actual de la propiedad o null si la propiedad no existe
     */
    @Override
    public String getProperty(String name) {
        CustomDocumentProperties custom = doc.getCustomDocumentProperties();

        if (custom.get(name) != null) {
            return custom.get(name).getValue().toString();
        }

        return null;
    }

    /**
     * Marca el documento como final, con el fin de prevenir la posterior
     * edicion del documento. Versiones mas recientes de Word permiten saltar
     * esta restriccion con un botï¿½n.
     */
    @Override
    public void markAsFinal() {
        CustomDocumentProperties custom = doc.getCustomDocumentProperties();

        if (custom.get("_MarkAsFinal") == null) {
            custom.add("_MarkAsFinal", true);
        }
    }

    /**
     * Si el documento esta marcado como final, elimina esta configuracion
     */
    @Override
    public void unmarkAsFinal() {
        CustomDocumentProperties custom = doc.getCustomDocumentProperties();

        if (custom.get("_MarkAsFinal") == null) {
            custom.add("_MarkAsFinal", false);
        } else {
            custom.get("_MarkAsFinal").setValue(false);
        }
    }

    /**
     * Protege este documento con contraseï¿½a. El tipo de protecciï¿½n es
     * "Rellenando Formularios".
     *
     * @param password El password que tendrï¿½ el documento protegido
     * 
     */
    @Override
    public void protect(String password) {
        doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS, password);
    }

    /**
     * Protege el documento con el mismo password que ya tenï¿½a. El tipo de
     * protecciï¿½n es "Rellenando Formularios".
     *
     * 
     */
    public void protect() {
        doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS);
    }

    /**
     * Protege este documento con contraseÃ±a. El tipo de protecciï¿½n es "Solo
     * Lectura".
     *
     * @param s El password que tendra el documento protegido
     * 
     */
    @Override
    public void protectReadOnly(String s) {
        doc.protect(ProtectionType.READ_ONLY, s);
    }

    /**
     * Elimina la proteccion de un documento sin necesidad de contraseï¿½a.
     *
     * 
     */
    @Override
    public void eraseProtection() {
        doc.protect(ProtectionType.NO_PROTECTION);
        unmarkAsFinal();
    }

    /**
     * Verifica si un documento esta protegido o no
     *
     * @return true si el documento esta protegido, de lo contrario false
     */
    @Override
    public boolean isProtected() {
        if (doc.getProtectionType() == ProtectionType.NO_PROTECTION) {
            return false;
        } else {
            return true;
        }

    }

    /**
     * Desprotege un documento suministrando su password
     *
     * @param password El password del documento
     * 
     */
    @Override
    public boolean unprotect(String password) {
    	boolean ret = doc.unprotect(password);
        logger.info("unprotect " + ret);
        return ret;
    }

    /**
     * Desprotege un documento sin necesidad de password
     *
     * 
     */
    @Override
    public void unprotectWithoutPassword() {
        doc.removeMacros();
        doc.unprotect();
    }

    @Override
    public List<String> validateBookmarkStructure() {
        ArrayList<String> mensajes = new ArrayList<>();

        byte[] file = getFile();

        try {
            File temp = File.createTempFile("blank_", ".docx");
            temp.createNewFile();

            FileUtils.writeByteArrayToFile(temp, file);
            try (ZipFile zip = new ZipFile(temp)) {
                Enumeration<? extends ZipEntry> entries = zip.entries();

                while (entries.hasMoreElements()) {
                    ZipEntry entry = entries.nextElement();
                    String entryName = entry.getName();
                    logger.info("Validando contenido: " + entryName);
                    if (!entryName.contains("document.xml") && !entryName.contains("header")
                            && !entryName.contains("footer")) {
                        logger.info(entryName + " no es valido");
                        continue;
                    }

                    byte[] content = IOUtils.toByteArray(zip.getInputStream(entry));

                    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                    factory.setNamespaceAware(true);
                    org.w3c.dom.Document xmlDocument = factory.newDocumentBuilder()
                            .parse(new ByteArrayInputStream(content));
                    String objectModel = NamespaceConstant.OBJECT_MODEL_SAXON;
                    System.setProperty("javax.xml.xpath.XPathFactory:" + objectModel,
                            "net.sf.saxon.xpath.XPathFactoryImpl");

                    XPathFactory xpathFactory = XPathFactory.newInstance(objectModel);
                    XPath xpath = xpathFactory.newXPath();

                    xpath.setNamespaceContext(new NamespaceContext() {
                        public String getNamespaceURI(String prefix) {
                            if (prefix == null) {
                                throw new NullPointerException("Null prefix");
                            } else if ("w".equals(prefix)) {
                                return "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                            } else if ("xml".equals(prefix)) {
                                return XMLConstants.XML_NS_URI;
                            }
                            return XMLConstants.NULL_NS_URI;
                        }

                        public String getPrefix(String uri) {
                            throw new UnsupportedOperationException();
                        }

                        public Iterator getPrefixes(String uri) {
                            throw new UnsupportedOperationException();
                        }
                    });

                    List<BookMarkProperties> marksNames = getMarksNames();

                    for (BookMarkProperties bk : marksNames) {

                        // Buscar el id de la marca
                        String exprLocateId = "//w:bookmarkStart[@w:name=\"" + bk.getName() + "\"]";
                        PlantillasLogger.mainLogger.info("Localizar el id de la marca : " + exprLocateId);
                        javax.xml.xpath.XPathExpression xpathLocate = xpath.compile(exprLocateId);
                        DOMNodeList bkStartNode = (DOMNodeList) xpathLocate.evaluate(xmlDocument,
                                XPathConstants.NODESET);

                        if (bkStartNode.getLength() == 0) {
                            continue;
                        }

                        String bkId = bkStartNode.item(0).getAttributes().getNamedItem("w:id").getTextContent();

                        String exprGen = "//*[(. >> (//w:bookmarkStart[@w:id=\"" + bkId + "\" and @w:name=\""
                                + bk.getName() + "\"])[1]) " + "and . << (//w:bookmarkEnd[@w:id=\"" + bkId + "\"])[1]]";

                        PlantillasLogger.mainLogger.info("Expresi\u00f3n para nodos intermedios : " + exprLocateId);

                        javax.xml.xpath.XPathExpression expr = xpath.compile(exprGen);

                        DOMNodeList list = (DOMNodeList) expr.evaluate(xmlDocument, XPathConstants.NODESET);
                        for (int i = 0; i < list.getLength(); i++) {
                            String currTag = list.item(i).getLocalName();
                            boolean exists = invalidTagsBk.keySet().contains(currTag);

                            String info = "";
                            // Si es un inicio de marca hay que decir que marca
                            // estÃ¡ adentro
                            if ("bookmarkStart".equals(currTag)) {
                                String currName = list.item(i).getAttributes().getNamedItem("w:name").getTextContent();
                                if (!"_GoBack".equals(currName)) {
                                    info = " ('" + currName + "')";
                                } else {
                                    continue;
                                }
                            }

                            if (exists) {
                                mensajes.add("La marca '" + bk.getName() + "' " + invalidTagsBk.get(currTag) + info);
                            }
                        }

                    }
                }
            }

            return mensajes;
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    private boolean checkBlankPassword() {
        protect();
        byte[] file = getFile();

        try {
            File temp = File.createTempFile("blank_", ".docx");
            temp.createNewFile();
            FileUtils.writeByteArrayToFile(temp, file);

            try (ZipFile zip = new ZipFile(temp)) {
                ZipEntry entry = zip.getEntry("word/settings.xml");
                byte[] content = IOUtils.toByteArray(zip.getInputStream(entry));

                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                javax.xml.parsers.DocumentBuilder builder = factory.newDocumentBuilder();
                org.w3c.dom.Document docu = builder.parse(new ByteArrayInputStream(content));

                temp.delete();

                NodeList protNode = docu.getElementsByTagName("w:documentProtection");
                if (protNode.getLength() >= 1) {
                    org.w3c.dom.Node docProt = protNode.item(0);
                    int attrSize = docProt.getAttributes().getLength();

                    if (attrSize <= 2) {
                        zip.close();
                        return true;
                    } else {
                        zip.close();
                        return false;
                    }
                }

            }

            return false;
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    /**
     * Valida que la contraseÃ±a de un documento sea correcta
     *
     * @param password El password a verificar
     * 
     */
    @Override
    public void validatePassword(String password) {
        if (password == null) {
            throw new PlantillasAPIException(9);
        }

        int protectionType = doc.getProtectionType();
        try {
            boolean estaEnBlanco = checkBlankPassword();

            if (estaEnBlanco && !password.isEmpty()) {
                if (PlantillasLogger.mainLogger.isDebugEnabled()) {
                    String tempPath = System.getProperty("java.io.tmpdir") + "/";
                    FileUtils.writeByteArrayToFile(new File(tempPath + System.currentTimeMillis() + ".docx"),
                            getFile());
                }
                throw new PlantillasAPIException(2, password);
            }

            if (password.isEmpty()) {
                logger.info("Password vacio");
                if (!estaEnBlanco) {
                    if (PlantillasLogger.mainLogger.isDebugEnabled()) {
                        String tempPath = System.getProperty("java.io.tmpdir") + "/";
                        FileUtils.writeByteArrayToFile(new File(tempPath + System.currentTimeMillis() + ".docx"),
                                getFile());
                    }
                    throw new PlantillasAPIException(2, "");
                }

                logger.info("Password si es vacio");

                return;
            }

            logger.info("Password Normal");

            boolean resultado = doc.unprotect(password);
            doc.protect(protectionType);

            if (!resultado) {
                if (PlantillasLogger.mainLogger.isDebugEnabled()) {
                    String tempPath = System.getProperty("java.io.tmpdir") + "/";
                    FileUtils.writeByteArrayToFile(new File(tempPath + System.currentTimeMillis() + ".docx"),
                            getFile());
                }
                throw new PlantillasAPIException(2, password);
            }
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    /**
     * Guarda los cambios del documento actual y devuelve el documento en PDF
     *
     * @return Un arreglo de bytes con el documento convertido a PDF
     * 
     */
    @Override
    public byte[] convertToPDF() {
        ByteArrayOutputStream pdf = new ByteArrayOutputStream();

        try {
            doc.save(pdf, SaveFormat.PDF);
            return pdf.toByteArray();
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    public byte[] convertToPlainText() throws Exception {
        ByteArrayOutputStream pdf = new ByteArrayOutputStream();

        doc.save(pdf, SaveFormat.TEXT);

        return pdf.toByteArray();
    }

    private void renameBookmarksInCell(Cell cell, String suffix) throws Exception {
        NodeCollection nodes = cell.getChildNodes(NodeType.BOOKMARK_START, true);

        Iterator<BookmarkStart> iterator = nodes.iterator();
        while (iterator.hasNext()) {
            Bookmark bk = iterator.next().getBookmark();
            bk.setName(bk.getName().concat(suffix));
        }
    }

    public void combineBlock(Map<String, Object> bookmarks) {
        if (bookmarks == null) {
            throw new PlantillasAPIException(3);
        }

        Set<String> keys = bookmarks.keySet();

        if (keys.isEmpty()) {
            throw new PlantillasAPIException(3);
        }

        try {
            String suffix = null;
            for (String key : keys) {
                suffix = addBookmarkBlock(key);
                if (suffix != null) {
                    break;
                }
            }

            if (suffix == null) {
                throw new PlantillasAPIException(4);
            }

            for (Entry<String, Object> bookmarkEntry : bookmarks.entrySet()) {
                String bkName = bookmarkEntry.getKey();
                Object bkValue = bookmarkEntry.getValue();
                BookMarkProperties prop = findBookmark(bkName);

                Bookmark bk = null;
                bk = doc.getRange().getBookmarks().get(bkName + suffix);
                if (bk == null) {
                    logger.warn("No  pudo  encontrar la marca : " + bkName + suffix);
                    continue;
                }
                if (bkValue instanceof byte[]) {
                    byte[] img = (byte[]) bkValue;
                    // bk = doc.getRange().getBookmarks().get(bkName + suffix);
                    bk.setText("");
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    builder.moveTo(bk.getBookmarkStart());
                    bk.remove();
                    builder.startBookmark(bkName + suffix);
                    if (prop != null) {
                        builder.insertImage(img, prop.getWidth(), prop.getHeight());
                    }
                    builder.endBookmark(bkName + suffix);
                } else {

                    bk.setText(bkValue.toString());
                }
            }
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    private String addBookmarkBlock(String baseBk) throws Exception {
        String block = getProperty("block");
        if (block == null) {
            setProperty("block", "ok");

            return "";
        }

        Bookmark bk = doc.getRange().getBookmarks().get(baseBk);
        if(bk ==null ) {
        	logger.warn("00: Error tipo marca : " + baseBk);
        	return null;
        }
        BookmarkStart bstart = bk.getBookmarkStart();
        if(bstart==null) {
        	logger.warn("01: Error tipo marca : " + baseBk);
        	return null;
        }
        Node ancestor = bstart.getAncestor(NodeType.TABLE);

        if (ancestor == null) {
            return null;
        }

        Table table = (Table) ancestor;

        // Obtener el numero de columnas basado en la primera fila
        Row firstRow = table.getFirstRow();
        int cols = firstRow.getCells().getCount();

        // Clonar la celda base
        Cell cell = table.getFirstRow().getFirstCell();

        Cell clonedCell = (Cell) cell.deepClone(true);

        Row lastRow = table.getLastRow();

        // Si es la primera fila, ya estan las celdas
        if (lastRow.equals(table.getFirstRow())) {
            CellCollection cells = lastRow.getCells();

            Iterator<Cell> iterator = cells.iterator();
            int pos = 0;
            int correctPos = 0;
            while (iterator.hasNext()) {
                Cell currentCell = iterator.next();
                if (currentCell.getText().length() == 1) {
                    correctPos = pos;
                    break;
                }

                pos++;
            }

            if (pos != cols) {
                int rowIndex = 0;
                int colIndex = pos;

                String index = "#" + rowIndex + "," + colIndex;
                renameBookmarksInCell(clonedCell, index);

                cells.removeAt(correctPos);
                cells.insert(correctPos, clonedCell);

                return index;
            }
        }

        int currentCols = lastRow.getCells().getCount();

        if (currentCols == cols) {
            Row clonedRow = (Row) table.getLastRow().deepClone(false);
            table.appendChild(clonedRow);
        }

        int rowIndex = table.getRows().getCount() - 1;
        int colIndex = table.getLastRow().getCells().getCount();

        String index = "#" + rowIndex + "," + colIndex;
        renameBookmarksInCell(clonedCell, index);

        table.getLastRow().appendChild(clonedCell);

        return index;
    }

    public void print() {
        doc.print();
    }

    /**
     * Especifica de que carpeta deben tomarse las fuentes que se usaran en este
     * documento. Este metodo se usa mas que todo en sistemas Linux para
     * especificar donde se encuentran las fuentes de Windows que usan los
     * documentos.
     *
     * @param fontsFolder La carpeta donde se encuentren las fuentes.
     */
    @Override
    public void setFontsFolder(String fontsFolder) {
        if (fontsFolder != null) {
            FontSettings.getDefaultInstance().setFontsFolder(fontsFolder, true);
        }
    }

    public byte[] signDocument(byte[] certBytes, String password) {
        X509Certificate cert = null;
        Security.addProvider(new BouncyCastleProvider());

        try {
            KeyStore ks = KeyStore.getInstance("PKCS12");
            try {
                ks.load(new ByteArrayInputStream(certBytes), password.toCharArray());
            } catch (Exception exc) {
                if (exc.getMessage() == null) {
                    throw exc;
                }

                if (exc.getMessage().contains("Given final block not properly padded")) {
                    throw new PlantillasAPIException(1);
                } else if (exc.getMessage().contains("by zero")) {
                    throw new PlantillasAPIException(2);
                } else {
                    throw exc;
                }
            }

            Enumeration en = ks.aliases();
            String alias = "";
            Vector vectaliases = new Vector();

            while (en.hasMoreElements()) {
                vectaliases.add(en.nextElement());
            }
            String[] aliases = (String[]) (vectaliases.toArray(new String[0]));
            for (int i = 0; i < aliases.length; i++) {
                if (ks.isKeyEntry(aliases[i])) {
                    alias = aliases[i];
                    break;
                }
            }

            cert = (X509Certificate) ks.getCertificate(alias);

            ByteArrayInputStream in = new ByteArrayInputStream(getFile());
            ByteArrayOutputStream out = new ByteArrayOutputStream();

            DigitalSignatureUtil.sign(in, out, CertificateHolder.create(certBytes, password), "", new Date());

            return out.toByteArray();
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    public static void main(String[] args) throws Exception {
        ConsoleAppender console = new ConsoleAppender();
        String PATTERN = "%d [%p|%c|%C{1}] %m%n";
        console.setLayout(new PatternLayout(PATTERN));
        console.setThreshold(Level.ALL);
        console.activateOptions();
        Logger.getRootLogger().addAppender(console);

        DocxTemplate template = new DocxTemplate();
        
        byte[] arc = FileUtils.readFileToByteArray(new File("C:/home/BORRE_MOD_2_VECES_ANTES.docx"));//plantilla.docx
      //  byte[] img = FileUtils.readFileToByteArray(new File("C:/bad/firmota.jpg"));
        template.load(arc);
        Map<String, Object> bookmarks = new HashMap<>();
        bookmarks.put("entregado_por", "XXX");
		//template.combineBlock(bookmarks );
        //template.combineMarkText("entregado_por", "que vaina buena");
        template.combineMarkHTML("html_report", "que vaina buena");
        template.getFile();
        //template.combineMarkImg("logoEntidad", IOUtils.toByteArray(new FileInputStream("c:\\home\\logoEntidad.png")));
        //template.addHeader();
        //template.combineMarkText("gdocs_numrad", "20203030633");
        //template.combineMarkDoc("word_ini", FileUtils.readFileToByteArray(new File("c:\\home\\header.docx")));
        String text="Carlos mario Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incidunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquid ex ea commodi consequat. Quis aute iure reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint obcaecat cupiditat non proident, sunt in culpa qui officia deserunt mollit anim id est laboru\n\n";
      StringBuilder sb = new StringBuilder();
         for(int x=0;x< 50;++x)
        	sb.append(text);
       // template.combineMarkText("dependencia", sb.toString());
        //template.combineMarkText("Rad_Respuesta", "202041120300001491, 202041120300001491, 202041120300001491");
        new FileOutputStream("c:\\home\\outsda.docx").write(template.getFile());
        
        /*HashMap<String, Object> block = new HashMap<>();
        block.put("dinamic_gdocs_cargofirma", "Gerente");
        block.put("dinamic_gdocs_nomfirma", "Hostias Hostias");
        block.put("dinamic_gdocs_firma", img);

        template.combineBlock(block);

        byte[] res = template.convertToPDF();
        template.close();

        FileUtils.writeByteArrayToFile(new File("C:/bad/firmado.pdf"), res);*/
    }

    /**
     * Escribe una marca de agua en el documento (no disponible con Aspose).
     *
     * @param texto El texto de la marca de agua
     * 
     */
    public void generateWatermark(String texto) {
        // Implementacion marca de agua
    }

    /**
     * Guarda los cambios del documento (no disponible con Aspose).
     *
     * 
     */
    public void save() {
        // Implementacion guardado de documento
    }

    /**
     * Cierra el objeto principal de Aspose.Words para liberar recursos
     *
     * 
     */
    public void close() {
        doc = null;
    }

    /**
     * Guarda el archivo en disco
     *
     * @param file El objeto File representando la ubicacion donde se va a
     *            guardar
     * 
     * @deprecated
     */
    @Deprecated
    public void saveAs(File file) {
        // Implementacion guardar como del documento
    }

    /**
     * Borra el archivo cargado
     *
     * 
     * @deprecated
     */
    @Deprecated
    public void delete() {
        // Implementacion borrado de documento
    }

    /**
     * Valida la version de Office del documento
     *
     * @param file El archivo a verificar
     * @return El mismo archivo cargado con un XML interno reestructurado
     * @deprecated
     */
    @Deprecated
    public byte[] validateOfficeVersion(byte[] file) {
        return file;
    }

    /**
     * Verifica la version de Office con la que fue editada el documento
     *
     * @param file
     * @return Siempre retorna 2007
     * @deprecated
     */
    @Deprecated
    public String getOfficeVersion(byte[] file) {
        return "2007";
    }

    @Override
    public int getPages() {
        try {
            doc.updatePageLayout();
            return doc.getPageCount();
        } catch (Exception e) {
            logger.error("Error general", e);
            throw new PlantillasAPIException(10, e);
        }
    }

    /**
     * Metodo para abrir una plantilla que tiene password
     * 
     * @param file
     * @param passwordPlantilla
     */
    @Override
    public void load(byte[] file, String passwordPlantilla) throws PlantillasAPIException {
        open(file, passwordPlantilla);
    }

    /**
     * Metodo sobrecargado que abre un archivo dado un arreglo de bytes que representa su contenido y su contrasena en
     * caso de tenerla
     * 
     * @param passwordPlantilla - Contrasena para abrir el documento
     * @param bytes El arreglo de bytes que contiene el archivo
     * 
     */
    public void open(byte[] bytes, String passwordPlantilla) throws PlantillasAPIException {
        /**
         * Tipo de error al verificar las contraseñas de los archivos protegidos.
         */
        int errorType = 0;

        try {
            License license = new License();
            InputStream lStream = this.getClass().getResourceAsStream("/Aspose.Total.Java.lic");
            logger.debug(lStream.toString());
            license.setLicense(lStream);
            if (!StringUtils.isBlank(passwordPlantilla)) {
                doc = new com.aspose.words.Document(new ByteArrayInputStream(bytes),
                        new LoadOptions(passwordPlantilla));
            } else {
                doc = new com.aspose.words.Document(new ByteArrayInputStream(bytes));
            }

            int protectionType = doc.getProtectionType();

            if (protectionType == 2) {
                if (StringUtils.isBlank(passwordPlantilla)) {
                    errorType = 11;
                    throw new PlantillasAPIException(11);
                } else {
                    boolean unprotect = doc.unprotect(passwordPlantilla);
                    System.out.println("Unprotect:  " + unprotect);
                    if (!unprotect) {
                        logger.info("Exception incorrect password");
                        errorType = 11;
                        throw new PlantillasAPIException(11);
                    }
                }
            } else if (protectionType == -1) {
                if (!StringUtils.isBlank(passwordPlantilla)) {
                    errorType = 12;
                    throw new PlantillasAPIException(11);
                }
            }

            logger.info("All was ok");
            obtenerMarcasDoc();

            // Cargar las marcas invalidas
            invalidTagsBk.put("tbl", "contiene una tabla");
            invalidTagsBk.put("bookmarkStart", "contiene un inicio de marca");
            invalidTagsBk.put("bookmarkEnd", "contiene un fin de marca");

            isOpen = true;

        } catch (Exception ex) {
            logger.error("Error general", ex);
            if (ex.getMessage().length() == 0 && ex.getMessage().contains("from ZipFile")) {
                throw new PlantillasAPIException(1);
            }
            if (ex.getMessage().contains("Unknown file format")) {
                throw new PlantillasAPIException(1);
            }
            if (ex.getMessage().contains("MP-PLT-0011")) {
                if (errorType == 12) {
                    throw new PlantillasAPIException(11,
                            "El documento seleccionado no tiene password, debe dejar el campo clave sin diligenciar");
                } else {
                    throw new PlantillasAPIException(11,
                            "El documento seleccionado tiene password y la clave digitada es incorrecta");
                }
            }
            throw new PlantillasAPIException(10, ex);
        }

    }
    
    
    /**
	 * Extrae una fraccion de imagen de la plantilla de un documento, de una seccion rectangular definida por los parametros
	 * @param x0 coordenada x unicial
	 * @param y0 coordenada y inicial
	 * @param width ancho del recuadro
	 * @param height alto del recuadro
	 * @return
     * @throws Exception 
	 * @throws IOException
	 */
	public byte[] extractImage(int x0, int y0, int width, int height) throws Exception {
		ByteArrayInputStream input = new ByteArrayInputStream(extractImage());
		BufferedImage bufferedImage0 = ImageIO.read(input);
		BufferedImage sectionImage = bufferedImage0.getSubimage(x0, y0, width, height);
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		ImageIO.write(sectionImage, "png", output);
		return output.toByteArray();
	} 
	
	/**
	 * Extrae una imagen de la plantilla de un documento, de la primera pagina
	 * @return
     * @throws Exception 
	 * @throws IOException
	 */
	public byte[] extractImage() throws Exception {
		ImageSaveOptions options = new ImageSaveOptions(SaveFormat.JPEG);
		options.setJpegQuality(100);
		options.setPageIndex(0);
		options.setPageCount(0);
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		doc.save(out, options);
		return out.toByteArray();
	}
	
	public void addHeader() throws Exception {
		System.out.println("Entro a esta verga");
		if(doc.getPageCount()>1) {
			System.out.println("Mas de una hoja");
		DocumentBuilder builder = new DocumentBuilder(doc);
		if(!builder.getPageSetup().getDifferentFirstPageHeaderFooter()) {
		builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);
		builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
		BufferedImage imageBytes = ImageIO.read(new File("c:\\home\\qr.png"));
		builder.insertImage(imageBytes );
		}
		}
	}
}
