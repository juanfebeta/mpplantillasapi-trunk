package com.macroproyectos.plantillas;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import com.macroproyectos.plantillas.errores.PlantillasAPIException;

/**
 * Interfaz que define los metodos con los cuales se puede manipular un
 * documento de Word
 *
 * @author CarlosMario
 */
public abstract interface DocumentTemplates {

    /**
     * Abre un archivo dada una ubicacion en disco
     * @param passwordDoc contraseña del documento en caso de tenerla
     * @param path La ruta donde se haya el archivo de Word
     * 
     */
    public abstract void open(String path);

    /**
     * Carga un archivo de Word para manipular con esta API.
     * @param passwordDoc contraseña del documento en caso de tenerla
     * @param file Un arreglo de bytes que representa el archivo a cargar
     * 
     */
    public abstract void load(byte[] file);

    /**
     * Guarda los cambios ejecutados en el archivo y devuelve un arreglo de
     * bytes con el archivo ya modificado.
     *
     * @return Un arreglo de bytes con el archivo ya modificado
     * 
     */
    public abstract byte[] getFile();

    /**
     * Toma todas las secciones del documento y las protege con el metodo
     * "Rellenando Formularios"
     */
    public abstract void lockAllSections();

    /**
     * Marca el documento como final, con el fin de prevenir la posterior
     * ediciï¿½n del documento. Versiones mas recientes de Word permiten saltar
     * esta restriccion con un botï¿½n.
     */
    public abstract void markAsFinal();

    /**
     * Habilita el control de cambios dentro del documento
     */
    public abstract void enableTrackChanges();

    /**
     * Acepta todos las revisiones del documento
     *
     * 
     */
    public abstract void acceptAllChanges();

    /**
     * Deshabilita el control de cambios en el documento
     */
    public abstract void disableTrackChanges();

    /**
     * Escribe una marca de agua en el documento (no disponible con Aspose).
     *
     * @param texto El texto de la marca de agua
     * 
     */
    public abstract void generateWatermark(String texto);

    /**
     * Borra todos los comentarios dentro de un documento
     *
     * 
     */
    public abstract void deleteAllComments();

    /**
     * Guarda los cambios del documento (no disponible con Aspose).
     *
     * 
     */
    public abstract void save();

    /**
     * Obtiene un arreglo de BookMarkProperties con todas las marcas que se
     * encuentren en el documento.
     *
     * @return null si el documento no esta abierto, de lo contrario, un
     * List de todas las marcas
     */
    public abstract List<BookMarkProperties> getMarksNames();

    /**
     * Asigna un texto a una marca en el documento
     *
     * @param marca El nombre de la marca
     * @param valor El contenido que se desea incluir en la marca
     * 
     */
    public abstract void combineMarkText(String marca, String valor);

    /**
     * Asigna contenido HTML a una marca en el documento
     *
     * @param marca El nombre de la marca
     * @param valor El contenido del documento HTML
     * 
     */
    public abstract void combineMarkHTML(String marca, String valor);

    /**
     * Asigna contenido HTML a una marca en el documento
     *
     * @param marca El nombre de la marca
     * @param valor El contenido del documento HTML
     * 
     */
    public abstract void combineMarkImg(String marca, byte[] valor);

    /**
     * Asigna otro documento de Word al contenido de una marca
     *
     * @param marca El nombre de la marca
     * @param valor El contenido del documento a incluir
     * 
     */
    public abstract void combineMarkDoc(String marca, byte[] valor);

    public abstract void combineBlock(Map<String, Object> bookmarks);

    /**
     * Protege este documento con contraseÃ±a. El tipo de protecciï¿½n es
     * "Rellenando Formularios".
     *
     * @param password El password que tendra el documento protegido
     * 
     */
    public abstract void protect(String password);

    /**
     * Protege este documento con contraseÃ±a. El tipo de protecciï¿½n es "Solo
     * Lectura".
     *
     * @param password El password que tendra el documento protegido
     * 
     */
    public abstract void protectReadOnly(String password);

    /**
     * Desprotege un documento suministrando su password
     *
     * @param password El password del documento
     * @return desprotegido
     */
    public abstract boolean unprotect(String password);

    /**
     * Elimina la protecciï¿½n de un documento sin necesidad de contraseï¿½a.
     *
     * 
     */
    public abstract void eraseProtection();

    /**
     * Cierra el objeto principal de Aspose.Words para liberar recursos
     *
     * 
     */
    public abstract void close();

    /**
     * Guarda el archivo en disco
     *
     * @param file El objeto File representando la ubicacion donde se va a
     * guardar
     * 
     * @deprecated
     */
    @Deprecated
    public abstract void saveAs(File file);

    /**
     * Borra el archivo cargado
     *
     * 
     * @deprecated
     */
    @Deprecated
    public abstract void delete();

    /**
     * Guarda los cambios del documento actual y devuelve el documento en PDF
     *
     * @return Un arreglo de bytes con el documento convertido a PDF
     * 
     */
    public abstract byte[] convertToPDF();

    /**
     * Recupera el valor actual de la versiï¿½n del documento
     *
     * @return El valor actual de la versiï¿½n
     */
    public abstract String getVersion();

    /**
     * Asigna al documento una version para propositos de verificacion interna
     *
     * @param ver El valor que se desea usar como versiï¿½n
     */
    public abstract void setVersion(String ver);

    /**
     * Recupera el valor de una propiedad
     *
     * @param name El nombre de la propiedad
     * @return El valor actual de la propiedad o null si la propiedad no existe
     */
    public abstract String getProperty(String name);

    /**
     * Crea una propiedad personalizada dentro del documento, si esta propiedad
     * ya existe, reemplaza su valor con el suministrado.
     *
     * @param name El nombre de la propiedad
     * @param val El valor de la propiedad
     */
    public abstract void setProperty(String name, String val);

    /**
     * Valida la version de Office del documento
     *
     * @param file El archivo a verificar
     * @return El mismo archivo cargado con un XML interno reestructurado
     * @deprecated
     */
    @Deprecated
    public abstract byte[] validateOfficeVersion(byte[] file);

    /**
     * Verifica la version de Office con la que fue editada el documento
     *
     * @param file
     * @return Siempre retorna 2007
     * @deprecated
     */
    @Deprecated
    public abstract String getOfficeVersion(byte[] file);

    /**
     * Valida que la contraseÃ±a de un documento sea correcta
     *
     * @param password El password a verificar
     * 
     */
    public abstract void validatePassword(String password);

    /**
     * Verifica si un documento esta protegido o no
     *
     * @return true si el documento esta protegido, de lo contrario false
     * 
     */
    public abstract boolean isProtected();
    
    public abstract List<String> validateBookmarkStructure();
    public abstract int getPages();
    
    public abstract void combineMarkGrid(String marca, List<String> headers, List<ArrayList> datagrid);
    
    public void unprotectWithoutPassword();
    
    public void unmarkAsFinal();
    
    public void setFontsFolder(String fontsFolder);
    
    public byte[] signDocument(byte[] certBytes, String password);
    
    /**
     * Carga un archivo de Word con password para manipular con esta API.
     * @param passwordDoc - contrasena del documento en caso de tenerla
     * @param file Un arreglo de bytes que representa el archivo a cargar
     * 
     */
    public abstract void load(byte[] file, String passwordPlantilla) throws PlantillasAPIException;

}
