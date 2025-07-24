// MenuManager.gs
class MenuManager {
  static buildMenu() {
    const ui = SpreadsheetApp.getUi();
    
    const menu = ui.createMenu('📊 Sistema Financiero')
      // Procesamiento principal
      .addItem('✉️ Procesar Correos', 'processEmails')
      .addSeparator()
      
      // Configuraciones
      .addSubMenu(ui.createMenu('⚙️ Configuración')
        .addItem('🔄 Actualizar Parámetros', 'refreshConfig')
        .addSeparator()
        .addSubMenu(ui.createMenu('📅 Rango Fechas')
          .addItem('Establecer Fecha Inicio', 'setStartDate')
          .addItem('Establecer Fecha Fin', 'setEndDate')
          .addItem('🔄 Restablecer Fechas', 'resetDates')
        )
        .addSeparator()
        .addSubMenu(ui.createMenu('🚫 Excepciones')
          .addItem('➕ Agregar Correo', 'addException')
          .addItem('➖ Quitar Correo', 'removeException')
          .addItem('📋 Listar Excepciones', 'listExceptions')
        )
      )
      
      // Acciones rápidas
      .addSeparator()
      .addSubMenu(ui.createMenu('⚡ Acciones Rápidas')
        .addItem('📤 Exportar Datos', 'exportData')
        .addItem('📊 Generar Reporte', 'generateReport')
      )
      
      // Ayuda y mantenimiento
      .addSeparator()
      .addItem('📝 Registro de Errores', 'showErrorLog')
      .addItem('ℹ️ Acerca del Sistema', 'showAbout')
      .addItem('🆘 Ayuda Rápida', 'showHelp');
      
    menu.addToUi();
  }

  // Función para forzar actualización del menú
  static refreshMenu() {
    this.buildMenu();
    SpreadsheetApp.getUi().alert('✅ Menú actualizado correctamente');
  }
}

// Función global para activar con onOpen (agregar en otro archivo .gs si es necesario)
function onOpen() {
  MenuManager.buildMenu();
}

// Ejemplo de función utilitaria adicional
function showHelp() {
  const helpContent = `
📌 **Uso del Sistema:**
1. Procesar Correos: Extrae datos de emails
2. Configuración: Personaliza fechas y excepciones
3. Acciones Rápidas: Exporta datos o genera reportes

✉️ *Soporte: hernandsayerm@gmail.com*`;
  
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(helpContent)
      .setWidth(400)
      .setHeight(300),
    '🆘 Ayuda Rápida'
  );
}
