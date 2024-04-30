import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.CellType;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Row;

// Definimos la clase CuentaBancaria
class CuentaBancaria {
    // Declaramos las variables de instancia para el titular, numero de cuenta y saldo
    private String titular;
    private int numeroCuenta;
    private double saldo;

    // Constructor de la clase CuentaBancaria
    public CuentaBancaria(String titular, int numeroCuenta, double saldo) {
        this.titular = titular;
        this.numeroCuenta = numeroCuenta;
        this.saldo = saldo;
    }

    // Metodo para depositar dinero en la cuenta
    public void depositar(double cantidad) {
        saldo += cantidad;
        System.out.printf("Se depositaron $%.2f en la cuenta de %s%n", cantidad, titular);
    }

    // Metodo para retirar dinero de la cuenta
    public void retirar(double cantidad) {
        if (saldo >= cantidad) {
            saldo -= cantidad;
            System.out.printf("Se retiraron $%.2f de la cuenta de %s%n", cantidad, titular);
        } else {
            System.out.println("Saldo insuficiente para realizar el retiro");
        }
    }

   // Metodo para obtener el saldo de la cuenta
public double getSaldo() {
    return saldo;
}

// Metodo para obtener el titular de la cuenta
public String getTitular() {
    return titular;
}

// Metodo para obtener el numero de la cuenta
public int getNumeroCuenta() {
    return numeroCuenta;
}

// Metodo para transferir dinero a otra cuenta
public boolean transferir(CuentaBancaria cuentaDestino, double cantidad) {
    // Verificamos si el saldo es suficiente para la transferencia
    if (saldo >= cantidad) {
        // Si es suficiente, restamos la cantidad del saldo
        saldo -= cantidad;
        // Depositamos la cantidad en la cuenta destino
        cuentaDestino.depositar(cantidad);
        // Imprimimos un mensaje indicando la cantidad transferida y las cuentas involucradas
        System.out.printf("Se transfirieron $%.2f de la cuenta de %s a la cuenta de %s%n", cantidad, titular, cuentaDestino.getTitular());
        // Devolvemos true para indicar que la transferencia fue exitosa
        return true;
    } else {
        // Si el saldo no es suficiente, imprimimos un mensaje de error
        System.out.println("Saldo insuficiente para realizar la transferencia");
        // Devolvemos false para indicar que la transferencia no fue exitosa
        return false;
    }
}

// Metodo para pagar un servicio
public boolean pagarServicio(String servicio, double cantidad) {
    // Verificamos si el saldo es suficiente para el pago
    if (saldo >= cantidad) {
        // Si es suficiente, restamos la cantidad del saldo
        saldo -= cantidad;
        // Imprimimos un mensaje indicando la cantidad pagada y el servicio
        System.out.printf("Se pagaron $%.2f por el servicio de %s%n", cantidad, servicio);
        // Devolvemos true para indicar que el pago fue exitoso
        return true;
    } else {
        // Si el saldo no es suficiente, imprimimos un mensaje de error
        System.out.println("Saldo insuficiente para realizar el pago");
        // Devolvemos false para indicar que el pago no fue exitoso
        return false;
    }
}

}

// Definimos la clase LectorExcel
class LectorExcel {
    // Metodo para leer cuentas desde un archivo Excel
    public static Map<Integer, CuentaBancaria> leerCuentasDesdeExcel(String rutaArchivo) throws IOException {
        // Creamos un mapa para almacenar las cuentas
        Map<Integer, CuentaBancaria> cuentas = new HashMap<>();
        // Abrimos el archivo Excel
        try (FileInputStream archivo = new FileInputStream(rutaArchivo);
             XSSFWorkbook libro = new XSSFWorkbook(archivo)) {
            // Obtenemos la primera hoja del libro
            XSSFSheet hoja = libro.getSheetAt(0);
            // Creamos un formateador para convertir las celdas a texto
            DataFormatter formatter = new DataFormatter();
            // Iteramos sobre cada fila de la hoja
            for (Row fila : hoja) {
                // Omitimos la primera fila que contiene los titulos de las columnas
                if (fila.getRowNum() == 0) {
                    continue;
                }
                // Obtenemos las celdas con el titular, numero de cuenta y saldo
                Cell celdaTitular = fila.getCell(0);
                Cell celdaNumero = fila.getCell(1);
                Cell celdaSaldo = fila.getCell(2);
                // Verificamos que ninguna de las celdas sea nula
                if (celdaTitular != null && celdaNumero != null && celdaSaldo != null) {
                    // Obtenemos los valores de las celdas
                    String titular = celdaTitular.getStringCellValue();
                    int numeroCuenta = (int) celdaNumero.getNumericCellValue();
                    double saldo = celdaSaldo.getNumericCellValue();
                    // Creamos una nueva cuenta bancaria y la añadimos al mapa
                    cuentas.put(numeroCuenta, new CuentaBancaria(titular, numeroCuenta, saldo));
                }
            }
        }
        // Devolvemos el mapa de cuentas
        return cuentas;
    }

    // Metodo para actualizar el saldo de una cuenta en un archivo Excel
public static void actualizarSaldoEnExcel(String rutaArchivo, int numeroCuenta, double nuevoSaldo) throws IOException {
    // Abrimos el archivo Excel
    try (FileInputStream archivo = new FileInputStream(rutaArchivo);
         XSSFWorkbook libro = new XSSFWorkbook(archivo)) {
        // Obtenemos la primera hoja del libro
        XSSFSheet hoja = libro.getSheetAt(0);
        // Iteramos sobre cada fila de la hoja
        for (Row fila : hoja) {
            // Omitimos la primera fila que contiene los titulos de las columnas
            if (fila.getRowNum() == 0) {
                continue;
            }
            // Obtenemos la celda con el numero de cuenta
            Cell celdaNumero = fila.getCell(1);
            // Obtenemos el valor de la celda
            int numeroCuentaFila = (int) celdaNumero.getNumericCellValue();
            // Verificamos si el numero de cuenta coincide con el que queremos actualizar
            if (numeroCuentaFila == numeroCuenta) {
                // Obtenemos la celda con el saldo
                Cell celdaSaldo = fila.getCell(2);
                // Si la celda no existe, la creamos
                if (celdaSaldo == null) {
                    celdaSaldo = fila.createCell(2, CellType.NUMERIC);
                }
                // Actualizamos el valor de la celda con el nuevo saldo
                celdaSaldo.setCellValue(nuevoSaldo);
                // Salimos del bucle ya que encontramos la cuenta que queriamos actualizar
                break;
            }
        }
        // Guardamos los cambios en el archivo Excel
        try (FileOutputStream salida = new FileOutputStream(rutaArchivo)) {
            libro.write(salida);
        }
    }
}

   // Metodo para agregar una nueva cuenta a un archivo Excel
public static void agregarCuentaAExcel(String rutaArchivo, CuentaBancaria nuevaCuenta) throws IOException {
    // Abrimos el archivo Excel
    try (FileInputStream archivo = new FileInputStream(rutaArchivo);
         XSSFWorkbook libro = new XSSFWorkbook(archivo)) {
        // Obtenemos la primera hoja del libro
        XSSFSheet hoja = libro.getSheetAt(0);
        // Obtenemos el numero de la ultima fila
        int ultimaFila = hoja.getLastRowNum();
        // Creamos una nueva fila al final
        Row fila = hoja.createRow(ultimaFila + 1);
        // Creamos las celdas para el titular, numero de cuenta y saldo
        fila.createCell(0).setCellValue(nuevaCuenta.getTitular());
        fila.createCell(1).setCellValue(nuevaCuenta.getNumeroCuenta());
        fila.createCell(2).setCellValue(nuevaCuenta.getSaldo());
        // Guardamos los cambios en el archivo Excel
        try (FileOutputStream salida = new FileOutputStream(rutaArchivo)) {
            libro.write(salida);
        }
    }
}
    
 // Metodo para eliminar una cuenta de un archivo Excel
public static void eliminarCuentaDeExcel(String rutaArchivo, int numeroCuenta) throws IOException {
    // Abrimos el archivo Excel
    try (FileInputStream archivo = new FileInputStream(rutaArchivo);
         XSSFWorkbook libro = new XSSFWorkbook(archivo)) {
        // Obtenemos la primera hoja del libro
        XSSFSheet hoja = libro.getSheetAt(0);
        // Inicializamos el numero de fila a 0
        int numeroFila = 0;
        // Iteramos sobre cada fila de la hoja
        for (Row fila : hoja) {
            // Omitimos la primera fila que contiene los titulos de las columnas
            if (fila.getRowNum() == 0) {
                continue;
            }
            // Obtenemos la celda con el numero de cuenta
            Cell celdaNumero = fila.getCell(1);
            // Obtenemos el valor de la celda
            int numeroCuentaFila = (int) celdaNumero.getNumericCellValue();
            // Verificamos si el numero de cuenta coincide con el que queremos eliminar
            if (numeroCuentaFila == numeroCuenta) {
                // Si coincide, guardamos el numero de fila y salimos del bucle
                numeroFila = fila.getRowNum();
                break;
            }
        }
        // Si encontramos la cuenta que queremos eliminar
        if (numeroFila > 0) {
            // Obtenemos la fila para eliminar
            Row filaParaEliminar = hoja.getRow(numeroFila);
            // Eliminamos la fila de la hoja
            hoja.removeRow(filaParaEliminar);
        }
        // Guardamos los cambios en el archivo Excel
        try (FileOutputStream salida = new FileOutputStream(rutaArchivo)) {
            libro.write(salida);
            }
        }
    }
}

// Clase principal del sistema bancario
public class SistemaBancario {
    public static void main(String[] args) {
        // Ruta del archivo Excel, asegurate de cambiar esto por la ruta real de tu archivo
        String rutaArchivoExcel = "C:\\Users\\lew\\Desktop\\CH100423\\Banco\\resources\\archivo.xlsx";
        try {
            // Leemos las cuentas desde el archivo Excel
            Map<Integer, CuentaBancaria> cuentas = LectorExcel.leerCuentasDesdeExcel(rutaArchivoExcel);
            // Creamos un objeto Scanner para leer la entrada del usuario
            Scanner scanner = new Scanner(System.in);
            // Ciclo principal del programa
            while (true) {
                // Mostramos el menu al usuario
                System.out.println("Por favor, selecciona una opción:");
                System.out.println("1. Depositar en una cuenta");
                System.out.println("2. Retirar de una cuenta");
                System.out.println("3. Ver saldo de una cuenta");
                System.out.println("4. Crear una nueva cuenta");
                System.out.println("5. Eliminar una cuenta");
                System.out.println("6. Transferir entre cuentas");
                System.out.println("7. Pagar un servicio");
                System.out.println("8. Salir");
                // Leemos la opcion seleccionada por el usuario
                int opcion = scanner.nextInt();
                // Si la opcion es 8, salimos del ciclo y terminamos el programa
                if (opcion == 8) {
                    break;
                }
                // Si la opcion no es 4 (crear una nueva cuenta), pedimos al usuario que introduzca el numero de cuenta
                if (opcion != 4) {
                    System.out.println("Por favor, introduce el número de cuenta:");
                    int numeroCuenta = scanner.nextInt();
                    // Obtenemos la cuenta del mapa de cuentas
                    CuentaBancaria cuenta = cuentas.get(numeroCuenta);
                    // Si la cuenta no existe, mostramos un mensaje de error y volvemos al inicio del ciclo
                    if (cuenta == null) {
                        System.out.println("La cuenta no existe.");
                        continue;
                    }
                    // Dependiendo de la opcion seleccionada, realizamos diferentes operaciones
                    switch (opcion) {
                        case 1: // Depositar en una cuenta
                            System.out.println("Por favor, introduce la cantidad a depositar:");
                            double cantidadDeposito = scanner.nextDouble();
                            cuenta.depositar(cantidadDeposito);
                            LectorExcel.actualizarSaldoEnExcel(rutaArchivoExcel, cuenta.getNumeroCuenta(), cuenta.getSaldo());
                            break;
                        case 2: // Retirar de una cuenta
                            System.out.println("Por favor, introduce la cantidad a retirar:");
                            double cantidadRetiro = scanner.nextDouble();
                            cuenta.retirar(cantidadRetiro);
                            LectorExcel.actualizarSaldoEnExcel(rutaArchivoExcel, cuenta.getNumeroCuenta(), cuenta.getSaldo());
                            break;
                        case 3: // Ver saldo de una cuenta
                            System.out.printf("El saldo de la cuenta es: $%.2f%n", cuenta.getSaldo());
                            break;
                        case 5: // Eliminar una cuenta
                            cuentas.remove(numeroCuenta);
                            LectorExcel.eliminarCuentaDeExcel(rutaArchivoExcel, numeroCuenta);
                            System.out.println("La cuenta ha sido eliminada.");
                            break;
                        case 6: // Transferir entre cuentas
                            System.out.println("Por favor, introduce el número de la cuenta destino:");
                            int numeroCuentaDestino = scanner.nextInt();
                            CuentaBancaria cuentaDestino = cuentas.get(numeroCuentaDestino);
                            if (cuentaDestino == null) {
                                System.out.println("La cuenta destino no existe.");
                                continue;
                            }
                            // Solicitamos al usuario que introduzca la cantidad a transferir
System.out.println("Por favor, introduce la cantidad a transferir:");
double cantidadTransferencia = scanner.nextDouble();
// Intentamos realizar la transferencia
if (cuenta.transferir(cuentaDestino, cantidadTransferencia)) {
    // Si la transferencia es exitosa, actualizamos los saldos en el archivo Excel
    LectorExcel.actualizarSaldoEnExcel(rutaArchivoExcel, cuenta.getNumeroCuenta(), cuenta.getSaldo());
    LectorExcel.actualizarSaldoEnExcel(rutaArchivoExcel, cuentaDestino.getNumeroCuenta(), cuentaDestino.getSaldo());
}
// Fin de la opcion de transferencia
break;
case 7:
// Inicio de la opcion de pago de servicios
System.out.println("Por favor, selecciona el servicio a pagar:");
System.out.println("1. Agua");
System.out.println("2. Telefono");
System.out.println("3. Luz");
System.out.println("4. Universidad");
int opcionServicio = scanner.nextInt();
String servicio;
// Dependiendo de la opcion seleccionada, establecemos el servicio correspondiente
switch (opcionServicio) {
    case 1:
        servicio = "Agua";
        break;
    case 2:
        servicio = "Telefono";
        break;
    case 3:
        servicio = "Luz";
        break;
    case 4:
        servicio = "Universidad";
        break;
    default:
        System.out.println("Opcion no valida.");
        continue;
}
// Solicitamos al usuario que introduzca la cantidad a pagar
System.out.println("Por favor, introduce la cantidad a pagar:");
double cantidadPago = scanner.nextDouble();
// Intentamos realizar el pago
if (cuenta.pagarServicio(servicio, cantidadPago)) {
    // Si el pago es exitoso, actualizamos el saldo en el archivo Excel
    LectorExcel.actualizarSaldoEnExcel(rutaArchivoExcel, cuenta.getNumeroCuenta(), cuenta.getSaldo());
}
 break;
// Si la opcion seleccionada no es valida, mostramos un mensaje de error
default:
    System.out.println("Opcion no valida.");
}
} else {
// Si la opcion seleccionada es 4 (crear una nueva cuenta)
System.out.println("Por favor, introduce el nombre del titular:");
scanner.nextLine(); // Consumir el salto de linea pendiente
String titular = scanner.nextLine();
System.out.println("Por favor, introduce el numero de cuenta:");
int numeroCuenta = scanner.nextInt();
System.out.println("Por favor, introduce el saldo inicial:");
double saldoInicial = scanner.nextDouble();
// Creamos una nueva cuenta bancaria y la añadimos al mapa de cuentas
CuentaBancaria nuevaCuenta = new CuentaBancaria(titular, numeroCuenta, saldoInicial);
cuentas.put(numeroCuenta, nuevaCuenta);
// Agregamos la nueva cuenta al archivo Excel
LectorExcel.agregarCuentaAExcel(rutaArchivoExcel, nuevaCuenta);
}
}
// Cerramos el scanner antes de terminar el programa
scanner.close();
} catch (FileNotFoundException e) {
// Si el archivo Excel no se encontro, mostramos un mensaje de error
System.out.println("El archivo Excel no se encontro: " + e.getMessage());
} catch (IOException e) {
// Si hubo un error al leer o escribir el archivo Excel, mostramos un mensaje de error
System.out.println("Error al leer o escribir el archivo Excel: " + e.getMessage());
}
}
}

