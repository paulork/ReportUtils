package br.com.paulork.reportutils;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileFilter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringReader;
import java.sql.Connection;
import java.sql.ResultSet;
import java.util.HashMap;
import java.util.List;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JRResultSetDataSource;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import net.sf.jasperreports.engine.export.ooxml.JRDocxExporter;
import net.sf.jasperreports.engine.export.ooxml.JRPptxExporter;
import net.sf.jasperreports.engine.export.ooxml.JRXlsxExporter;
import net.sf.jasperreports.engine.query.JRXPathQueryExecuterFactory;
import net.sf.jasperreports.engine.util.JRLoader;
import net.sf.jasperreports.export.SimpleDocxReportConfiguration;
import net.sf.jasperreports.export.SimpleExporterInput;
import net.sf.jasperreports.export.SimpleOutputStreamExporterOutput;
import net.sf.jasperreports.export.SimplePptxReportConfiguration;
import net.sf.jasperreports.export.SimpleXlsxReportConfiguration;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

/**
 * @author Paulo R. Kraemer <paulork10@gmail.com>
 */
public class ReportUtils {

    /**
     * A partir do nome do relatório e de uma lista (da entidade) que representa
     * os dados, monta o relatório.
     *
     * @param relatorio Nome do relatório
     * @param datasource Lista da entidade requerida no relatório.
     * @return Retorna a classe JasperPrint que em seguida poderá ser convertida
     * para um array de bytes e salva em disco ou retransmitida
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public JasperPrint populaRelatorio(String relatorio, List<?> datasource) throws JRException {
        JRBeanCollectionDataSource ds = new JRBeanCollectionDataSource(datasource);
        JasperReport rel = getRelatorio(relatorio);
        return JasperFillManager.fillReport(rel, new HashMap(), ds);
    }

    /**
     * A partir do nome do relatório e de uma lista (da entidade) que representa
     * os dados, monta o relatório.
     *
     * @param relatorio Nome do relatório
     * @param datasource Lista da entidade requerida no relatório.
     * @param parametros Lista de parametros requeridos no relatório.
     * @return Retorna a classe JasperPrint que em seguida poderá ser convertida
     * para um array de bytes e salva em disco ou retransmitida
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public JasperPrint populaRelatorio(String relatorio, List<?> datasource, HashMap<String, Object> parametros) throws JRException {
        JRBeanCollectionDataSource ds = new JRBeanCollectionDataSource(datasource);
        JasperReport rel = getRelatorio(relatorio);
        return JasperFillManager.fillReport(rel, parametros, ds);
    }

    /**
     * A partir do nome do relatório e de um ResultSet populado com os dados
     * necessários, monta o relatório.
     *
     * @param relatorio Nome do relatório
     * @param datasource ResultSet com os dados necessários.
     * @return Retorna a classe JasperPrint que em seguida poderá ser convertida
     * para um array de bytes e salva em disco ou retransmitida
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public JasperPrint populaRelatorio(String relatorio, ResultSet datasource) throws JRException {
        JRResultSetDataSource ds = new JRResultSetDataSource(datasource);
        JasperReport rel = getRelatorio(relatorio);
        return JasperFillManager.fillReport(rel, new HashMap(), ds);
    }

    /**
     * A partir do nome do relatório e de uma representação em memória de um XML
     * (Document), monta o relatório.
     *
     * @param relatorio Nome do relatório
     * @param datasource Document representando o XML com os dados
     * @return Retorna a classe JasperPrint que em seguida poderá ser convertida
     * para um array de bytes e salva em disco ou retransmitida
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public JasperPrint populaRelatorio(String relatorio, Document datasource) throws JRException {
        HashMap params = new HashMap();
        params.put(JRXPathQueryExecuterFactory.PARAMETER_XML_DATA_DOCUMENT, datasource);
        return JasperFillManager.fillReport(getRelatorio(relatorio), params);
    }

    /**
     * A partir do nome do relatório e de uma conexão (Connection) com a base de
     * dados.
     *
     * @param relatorio Nome do relatório
     * @param datasource Conexão (geralmente Hibernate) com a base de dados
     * @return Retorna a classe JasperPrint que em seguida poderá ser convertida
     * para um array de bytes e salva em disco ou retransmitida
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public JasperPrint populaRelatorio(String relatorio, Connection datasource) throws JRException {
        JasperReport rel = getRelatorio(relatorio);
        return JasperFillManager.fillReport(rel, new HashMap(), datasource);
    }

    /**
     * A partir do nome do relatório e de um HashMap (esquema chave/valor) com
     * os dados que se quer representar no relatório. Geralmente usado para
     * simples listagem.
     *
     * @param relatorio Nome do relatório
     * @param datasource HashMap contendo os dados a serem representados no
     * relatório
     * @return Retorna a classe JasperPrint que em seguida poderá ser convertida
     * para um array de bytes e salva em disco ou retransmitida
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public JasperPrint populaRelatorio(String relatorio, HashMap<String, Object> datasource) throws JRException {
        JasperReport rel = getRelatorio(relatorio);
        return JasperFillManager.fillReport(rel, datasource);
    }

    /**
     * Carrega o modelo (layout) do relatório para que posteriormente possa se
     * populado com os dados provinientes do DataSource.
     *
     * @param caminhoRel Caminho do relatório
     * @return Modelo do relatório pronto para ser populados com os dados
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public JasperReport getRelatorio(String caminhoRel) throws JRException {
        if (caminhoRel.endsWith(".jasper")) {
            return (JasperReport) JRLoader.loadObjectFromFile(caminhoRel);
        } else if (caminhoRel.endsWith(".jrxml")) {
            return JasperCompileManager.compileReport(caminhoRel);
        } else {
            throw new JRException("Extensão de arquivo não reconhecida/permitida!");
        }
    }

    /**
     * Converte um relatório já populado com os dados, em um array de bytes.
     *
     * @param relatorio Relatório Jasper populado
     * @return Um array de bytes representando o arquivo PDF.
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public byte[] jasperToByte(JasperPrint relatorio) throws JRException {
        return JasperExportManager.exportReportToPdf(relatorio);
    }

    /**
     * Converte um relatório já populado com os dados para um Array de Byte Pode
     * posteriormente ser retornado pelo Servlet para exibição na web ou
     * simplesmente ser saldo no disco.
     *
     * @param relatorio Relatório Jasper populado
     * @return Uma string XML
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public String jasperToXml(JasperPrint relatorio) throws JRException {
        return JasperExportManager.exportReportToXml(relatorio);
    }

    /**
     * Converte um relatório já populado com os dados, em um array de bytes.
     *
     * @param relatorio Relatório Jasper populado
     * @return Um array de bytes representando o arquivo XLS.
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public byte[] jasperToXls(JasperPrint relatorio) throws JRException {
        JRXlsxExporter exporter = new JRXlsxExporter();
        exporter.setExporterInput(new SimpleExporterInput(relatorio));
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(baos));
        SimpleXlsxReportConfiguration configuration = new SimpleXlsxReportConfiguration();
//        configuration.setOnePagePerSheet(true);
//        configuration.setDetectCellType(true);
//        configuration.setCollapseRowSpan(false);
        exporter.setConfiguration(configuration);
        exporter.exportReport();
        return baos.toByteArray();
    }

    /**
     * Converte um relatório já populado com os dados, em um array de bytes.
     *
     * @param relatorio Relatório Jasper populado
     * @return Um array de bytes representando o arquivo DOC.
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public byte[] jasperToDoc(JasperPrint relatorio) throws JRException {
        JRDocxExporter exporter = new JRDocxExporter();
        exporter.setExporterInput(new SimpleExporterInput(relatorio));
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(baos));
        SimpleDocxReportConfiguration configuration = new SimpleDocxReportConfiguration();
        exporter.setConfiguration(configuration);
        exporter.exportReport();
        return baos.toByteArray();
    }

    /**
     * Converte um relatório já populado com os dados em um array de bytes.
     *
     * @param relatorio Relatório Jasper populado
     * @return Um array de bytes representando o arquivo PPT.
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public byte[] jasperToPpt(JasperPrint relatorio) throws JRException {
        JRPptxExporter exporter = new JRPptxExporter();
        exporter.setExporterInput(new SimpleExporterInput(relatorio));
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(baos));
        SimplePptxReportConfiguration configuration = new SimplePptxReportConfiguration();
        exporter.setConfiguration(configuration);
        exporter.exportReport();
        return baos.toByteArray();
    }

    /**
     * Converte um relatório já populado com os dados para um
     * ByteArrayOutputStream. Pode posteriormente ser retornado pelo Servlet
     * para exibição na web ou simplesmente ser saldo no disco.
     *
     * @param relatorio Relatório Jasper populado
     * @return Retorna um ByteArrayOutputStream
     * @throws JRException Caso ocorra algum problema na geração do relatório.
     */
    public ByteArrayOutputStream jasperToPdfStream(JasperPrint relatorio) throws JRException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        JasperExportManager.exportReportToPdfStream(relatorio, os);
        return os;
    }

    /**
     * Salva um Streaming de bytes puros para um arquivo pdf de destino.
     *
     * @param pdf Array de bytes representando o arquivo pdf.
     * @param destino Caminho de destino para o arquivo PDF.
     */
    public void savePdf(byte[] pdf, String destino) {
        try {
            FileOutputStream fos = new FileOutputStream(destino);
            fos.write(pdf);
            fos.flush();
            fos.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    /**
     * Salva um Streaming de bytes puros para um arquivo de destino.
     *
     * @param data Array de bytes representando o arquivo (PDF, XLS, HTML, etc).
     * @param destino Caminho de destino para o arquivo (PDF, XLS, HTML, etc).
     */
    public void saveFile(byte[] data, String destino) {
        savePdf(data, destino);
    }

    /**
     * Faz a transformação de uma string representando um XML para uma estrutura
     * de arvore XML (Document).
     *
     * @param xml String representando um XML
     * @return Document com a representação do XML (arvore de XML)
     * @throws ParserConfigurationException
     * @throws SAXException
     * @throws IOException
     */
    private Document strToDoc(String xml) throws Exception {
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            return builder.parse(new InputSource(new StringReader(xml)));
        } catch (ParserConfigurationException ex) {
            throw new ParserConfigurationException("Mensagem: " + ex.getMessage());
        } catch (SAXException ex) {
            throw new SAXException("Erro ao fazer o parser do XML. Mensagem: " + ex.getMessage());
        } catch (IOException ex) {
            throw new IOException("Erro de I/O. Mensagem: " + ex.getMessage());
        }
    }

    /**
     * Recompila o arquivo JRXML fornecido para JASPER.
     *
     * @param arquivoJRXML Caminho do arquivo
     * @throws JRException Erro ao compilar arquivo
     */
    private void recompila(String arquivoJRXML) throws JRException {
        JasperCompileManager.compileReportToFile(arquivoJRXML);
    }

    /**
     * Recompila todos os arquivo JRXML para JASPER do caminho fornecido.
     *
     * @param pastaJRXML Caminho da pasta
     * @throws JRException Erro ao compilar arquivo
     */
    private void recompilaTudo(String pastaJRXML) throws JRException {
        File[] files = listarArquivos(pastaJRXML);
        for (File file : files) {
            JasperCompileManager.compileReportToFile(file.getAbsolutePath());
        }
    }

    /**
     * Lista os arquivos JRXML da pasta fornecida por parametro.
     *
     * @param caminho Caminho da pasta
     * @return Um array de Files dos arquivos encontrados
     */
    private File[] listarArquivos(String caminho) {
        return listarArquivos(caminho, null);
    }

    /**
     * Listab arquivos de uma pasta (fornecida) de acordo com uma extenção
     * (fornecida).
     *
     * @param caminho Caminho da pasta
     * @param extencao Extenção dos arquivos a listar
     * @return Um array de Files dos arquivos encontrados
     */
    private File[] listarArquivos(String caminho, String extencao) {
        File fileDir = new File(caminho);
        final String ext;
        if (extencao == null) {
            ext = ".jrxml";
        } else {
            ext = (extencao.contains(".") ? extencao : "." + extencao);
        }

        // Listar todos os arquivos do diretório, com a dada extensão
        File[] listFiles = fileDir.listFiles(
            new FileFilter() {
                public boolean accept(File b) {
                    return b.getName().endsWith(ext);
                }
            }
        );
        return listFiles;
    }

}
