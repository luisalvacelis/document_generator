package dg.natos.mp;

import dg.natos.mp.controller.MainController;
import javax.swing.SwingUtilities;
import org.springframework.boot.WebApplicationType;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.context.ConfigurableApplicationContext;

@SpringBootApplication(scanBasePackages = "dg.natos.mp")
public class MainApp {

    public static void main(String[] args) {

        ConfigurableApplicationContext applicationContext = new SpringApplicationBuilder(MainApp.class)
                .headless(false)
                .web(WebApplicationType.NONE)
                .run(args);

        SwingUtilities.invokeLater(() -> {
            FileClass fileClass = new FileClass();
            MainController mainController = applicationContext.getBean(MainController.class);
            mainController.setFileClass(fileClass);
            mainController.setVisible(true);
        });
        //SpringApplication.run(MainApp.class, args);
    }
}
