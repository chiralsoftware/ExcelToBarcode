package chiralsoftware.exceltobarcode;

import java.util.logging.Logger;
import org.springframework.web.servlet.support.AbstractAnnotationConfigDispatcherServletInitializer;

/**
 *
 */
public class ServletConfiguration extends AbstractAnnotationConfigDispatcherServletInitializer {

    private static final Logger LOG = Logger.getLogger(ServletConfiguration.class.getName());

    @Override
    protected Class<?>[] getRootConfigClasses() {
        LOG.info("Getting root config classes");
        return new Class<?>[]{};
    }

    @Override
    protected Class<?>[] getServletConfigClasses() {
        LOG.info("Getting servlet config classes");
        return new Class<?>[]{WebConfiguration.class
        };
    }

    @Override
    protected String[] getServletMappings() {
        return new String[]{"/"};
    }

}
