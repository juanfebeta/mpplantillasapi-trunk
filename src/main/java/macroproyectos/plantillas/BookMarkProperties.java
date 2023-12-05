package com.macroproyectos.plantillas;

import java.io.Serializable;

/**
 * Bean para definir las propiedades de una marca de Word
 *
 * @author CarlosMario
 */
public class BookMarkProperties implements Serializable {

    public BookMarkProperties() {
    	//Constructor BookMarkProperties
    }

    /**
     * Obtiene el nombre de la marca.
     *
     * @return El nombre de la marca
     */
    public String getName() {
        return name;
    }

    /**
     * Asigna el nombre de la marca
     *
     * @param name El nombre de la marca
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * Obtiene el tipo de la marca (text o img)
     *
     * @return El tipo de la marca (text o img)
     */
    public String getType() {
        return type;
    }

    /**
     * Asigna el tipo de marca (text o img)
     *
     * @param type El tipo de marca (text o img)
     */
    public void setType(String type) {
        this.type = type;
    }

    /**
     * Retorna una cadena que contiene el nombre seguido de -- y luego el tipo
     * de la marca
     *
     * @return El nombre de la marca concatenado con -- y el tipo de la misma
     */
    public String toString() {
        return name + " -- " + type;
    }

    /**
     * Nombre de la marca
     */
    private String name;
    /**
     * Tipo de la marca
     */
    private String type;

    private String idTable;

    /**
     * @return the idTable
     */
    public String getIdTable() {
        return idTable;
    }

    /**
     * @param idTable the idTable to set
     */
    public void setIdTable(String idTable) {
        this.idTable = idTable;
    }
    
    private double height;
    private double width;
    
    public void setHeight(double height){
        this.height = height;
    }
    
    public double getHeight(){
        return height;
    }
    
    public void setWidth(double width){
        this.width = width;
    }
    
    public double getWidth(){
        return width;
    }
    

}
