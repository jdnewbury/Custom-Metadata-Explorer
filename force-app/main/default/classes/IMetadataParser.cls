/**
 * @description Interface for metadata parsers that convert XML content into structured data
 *              for display in the Custom Metadata Explorer.
 */
public interface IMetadataParser {
    
    /**
     * @description Parse XML content into a list of structured data rows
     * @param xmlContent Raw XML content from Bitbucket
     * @param filePath Path to the file in Bitbucket
     * @param config Configuration for this metadata type
     * @return List of maps representing rows of data
     */
    List<Map<String, Object>> parse(String xmlContent, String filePath, MetadataTypeConfig.TypeConfig config);
}
