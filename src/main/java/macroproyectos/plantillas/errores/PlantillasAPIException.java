package com.macroproyectos.plantillas.errores;

import com.mp.errors.MPException;

public class PlantillasAPIException extends MPException{
    
    private static final long serialVersionUID = 3852822835206545005L;

    public PlantillasAPIException(int number, Object... params) {
        this(number, null, params);
    }

    public PlantillasAPIException(int number, Exception cause, Object... params) {
        super(PlantillasAPIConstants.ERRORS_APP_MP, PlantillasAPIConstants.ERRORS_MP_PLT, number, cause, params);
    }
    
}
