
import javax.swing.JOptionPane;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author WIN8
 */
public class VentanaError {
    
    public static void mostarError( final String error ) {
        JOptionPane.showMessageDialog(null, error, "Error", JOptionPane.ERROR_MESSAGE);
    }
    
    public static void mostarSucces( final String error ) {
        JOptionPane.showMessageDialog(null, error, "Realizado", JOptionPane.INFORMATION_MESSAGE);
    }
    
}
