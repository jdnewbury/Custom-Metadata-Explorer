/**
 * @description Factory class that returns parsers for metadata types.
 *              Supports both config-driven parsing (using Custom Metadata column definitions)
 *              and specialized parsers for complex XML structures.
 *
 *              For most metadata types, the ConfigDrivenParser will work automatically
 *              based on the column XML paths defined in Bitbucket_Metadata_Column__mdt.
 *
 *              For complex types requiring special logic, create a named parser class
 *              and reference it in the Parser_Class__c field of Bitbucket_Metadata_Type__mdt.
 *
 * @author Jason Newbury
 * @date 2025
 */
public class MetadataParserFactory {

    /**
     * @description Get the appropriate parser for a metadata type
     * @param metadataType Developer name of the metadata type
     * @return IMetadataParser implementation
     */
    public static IMetadataParser getParser(String metadataType) {
        // Get config to check for custom parser class
        MetadataTypeConfig.TypeConfig config = MetadataTypeConfig.getConfig(metadataType);

        // If a custom parser class is specified, try to instantiate it
        if (String.isNotBlank(config.parserClass)) {
            IMetadataParser customParser = getCustomParser(config.parserClass);
            if (customParser != null) {
                return customParser;
            }
        }

        // Default to config-driven parser
        return new ConfigDrivenParser();
    }

    /**
     * @description Try to instantiate a custom parser by class name
     */
    private static IMetadataParser getCustomParser(String className) {
        // Map of known custom parsers
        // Add new specialized parsers here
        switch on className.toLowerCase() {
            when 'assignmentruleparser' { return new AssignmentRuleParser(); }
            when 'workflowparser' { return new WorkflowParser(); }
            when 'validationruleparser' { return new ValidationRuleParser(); }
            when 'approvalprocessparser' { return new ApprovalProcessParser(); }
            when 'flowparser' { return new FlowParser(); }
            when 'permissionsetparser' { return new PermissionSetParser(); }
            when 'profileparser' { return new ProfileParser(); }
            when 'sharingruleparser' { return new SharingRuleParser(); }
            when else {
                System.debug('Custom parser not found: ' + className + '. Using ConfigDrivenParser.');
                return null;
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // HELPER METHODS
    // ═══════════════════════════════════════════════════════════════════════

    public static String extractObjectName(String filePath) {
        String fileName = filePath.substringAfterLast('/');
        return fileName.substringBefore('.');
    }

    public static String getNodeText(Dom.XmlNode parent, String childName) {
        if (parent == null) return '';
        for (Dom.XmlNode child : parent.getChildElements()) {
            if (child.getName() == childName) {
                return child.getText() != null ? child.getText() : '';
            }
        }
        return '';
    }

    public static List<Dom.XmlNode> getChildNodes(Dom.XmlNode parent, String childName) {
        List<Dom.XmlNode> results = new List<Dom.XmlNode>();
        if (parent == null) return results;
        for (Dom.XmlNode child : parent.getChildElements()) {
            if (child.getName() == childName) {
                results.add(child);
            }
        }
        return results;
    }

    public static List<String> parseCriteriaItems(Dom.XmlNode parent) {
        List<String> criteria = new List<String>();
        for (Dom.XmlNode item : getChildNodes(parent, 'criteriaItems')) {
            String field = getNodeText(item, 'field');
            String operation = getNodeText(item, 'operation');
            String value = getNodeText(item, 'value');
            criteria.add(field + ' ' + operation + ' ' + value);
        }
        return criteria;
    }

    /**
     * @description Extract value from XML node using path notation
     * @param node XML node to search within
     * @param xmlPath Path like "fieldName" or "parent.child" or "$OBJECT_NAME"
     * @param filePath File path (for $OBJECT_NAME and $FILE_PATH)
     * @param isCollection Whether to collect multiple values
     * @param separator Separator for collections
     * @return Extracted value
     */
    public static String extractValueByPath(Dom.XmlNode node, String xmlPath, String filePath, 
                                            Boolean isCollection, String separator) {
        if (String.isBlank(xmlPath)) return '';

        // Handle special variables
        if (xmlPath == '$OBJECT_NAME') {
            return extractObjectName(filePath);
        }
        if (xmlPath == '$FILE_PATH') {
            return filePath;
        }
        if (xmlPath == '$FILE_NAME') {
            return filePath.substringAfterLast('/');
        }

        // Handle dot notation for nested paths
        List<String> pathParts = xmlPath.split('\\.');

        if (isCollection) {
            return extractCollectionValue(node, pathParts, separator);
        } else {
            return extractSingleValue(node, pathParts);
        }
    }

    private static String extractSingleValue(Dom.XmlNode node, List<String> pathParts) {
        Dom.XmlNode current = node;

        for (Integer i = 0; i < pathParts.size() - 1; i++) {
            List<Dom.XmlNode> children = getChildNodes(current, pathParts[i]);
            if (children.isEmpty()) return '';
            current = children[0];
        }

        return getNodeText(current, pathParts[pathParts.size() - 1]);
    }

    private static String extractCollectionValue(Dom.XmlNode node, List<String> pathParts, String separator) {
        List<String> values = new List<String>();
        String elementName = pathParts[0];

        for (Dom.XmlNode child : getChildNodes(node, elementName)) {
            if (pathParts.size() == 1) {
                String text = child.getText();
                if (String.isNotBlank(text)) values.add(text);
            } else {
                // Navigate nested path
                List<String> remainingPath = new List<String>();
                for (Integer i = 1; i < pathParts.size(); i++) {
                    remainingPath.add(pathParts[i]);
                }
                String val = extractSingleValue(child, remainingPath);
                if (String.isNotBlank(val)) values.add(val);
            }
        }

        return String.join(values, separator);
    }

    // ═══════════════════════════════════════════════════════════════════════
    // CONFIG-DRIVEN PARSER (Uses Custom Metadata column definitions)
    // ═══════════════════════════════════════════════════════════════════════
    public class ConfigDrivenParser implements IMetadataParser {

        public List<Map<String, Object>> parse(String xmlContent, String filePath, 
                                               MetadataTypeConfig.TypeConfig config) {
            List<Map<String, Object>> results = new List<Map<String, Object>>();

            Dom.Document doc = new Dom.Document();
            doc.load(xmlContent);
            Dom.XmlNode root = doc.getRootElement();

            // Determine what nodes to iterate over
            List<Dom.XmlNode> entryNodes = new List<Dom.XmlNode>();

            if (String.isNotBlank(config.entryElement)) {
                // If entry element specified, find all matching nodes
                if (String.isNotBlank(config.rootElement)) {
                    // First find root elements, then entry elements within
                    for (Dom.XmlNode rootNode : getChildNodes(root, config.rootElement)) {
                        entryNodes.addAll(getChildNodes(rootNode, config.entryElement));
                    }
                    // If no nested entries found, try entry elements directly under root
                    if (entryNodes.isEmpty()) {
                        entryNodes = getChildNodes(root, config.entryElement);
                    }
                } else {
                    entryNodes = getChildNodes(root, config.entryElement);
                }
            } else if (String.isNotBlank(config.rootElement)) {
                // Just root element, no nested entries
                entryNodes = getChildNodes(root, config.rootElement);
            } else {
                // No specific elements - treat root as single entry
                entryNodes.add(root);
            }

            // If still no entries, treat file as single record
            if (entryNodes.isEmpty()) {
                entryNodes.add(root);
            }

            // Parse each entry node
            for (Dom.XmlNode entryNode : entryNodes) {
                Map<String, Object> row = new Map<String, Object>();

                // Extract each column's value
                for (MetadataTypeConfig.ColumnConfig col : config.columns) {
                    if (!col.active) continue;

                    String value = extractValueByPath(
                        entryNode, 
                        col.xmlPath, 
                        filePath,
                        col.isCollection,
                        col.collectionSeparator
                    );

                    row.put(col.fieldName, value);
                }

                results.add(row);
            }

            return results;
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // SPECIALIZED PARSERS (For complex XML structures)
    // These are used when Parser_Class__c is set in Bitbucket_Metadata_Type__mdt
    // ═══════════════════════════════════════════════════════════════════════

    public class AssignmentRuleParser implements IMetadataParser {
        public List<Map<String, Object>> parse(String xmlContent, String filePath, 
                                               MetadataTypeConfig.TypeConfig config) {
            List<Map<String, Object>> results = new List<Map<String, Object>>();
            String objectName = extractObjectName(filePath);

            Dom.Document doc = new Dom.Document();
            doc.load(xmlContent);
            Dom.XmlNode root = doc.getRootElement();

            for (Dom.XmlNode assignmentRule : getChildNodes(root, 'assignmentRule')) {
                String fullName = getNodeText(assignmentRule, 'fullName');
                String active = getNodeText(assignmentRule, 'active');

                for (Dom.XmlNode ruleEntry : getChildNodes(assignmentRule, 'ruleEntry')) {
                    Map<String, Object> row = new Map<String, Object>();
                    row.put('objectName', objectName);
                    row.put('fullName', fullName);
                    row.put('active', active);
                    row.put('assignedTo', getNodeText(ruleEntry, 'assignedTo'));
                    row.put('assignedToType', getNodeText(ruleEntry, 'assignedToType'));
                    row.put('emailTemplate', getNodeText(ruleEntry, 'template'));

                    List<String> criteria = parseCriteriaItems(ruleEntry);
                    for (Integer i = 0; i < Math.min(criteria.size(), 7); i++) {
                        row.put('criteria' + (i + 1), criteria[i]);
                    }

                    results.add(row);
                }
            }
            return results;
        }
    }

    public class WorkflowParser implements IMetadataParser {
        public List<Map<String, Object>> parse(String xmlContent, String filePath, 
                                               MetadataTypeConfig.TypeConfig config) {
            List<Map<String, Object>> results = new List<Map<String, Object>>();
            String objectName = extractObjectName(filePath);

            Dom.Document doc = new Dom.Document();
            doc.load(xmlContent);
            Dom.XmlNode root = doc.getRootElement();

            for (Dom.XmlNode rule : getChildNodes(root, 'rules')) {
                Map<String, Object> row = new Map<String, Object>();
                row.put('objectName', objectName);
                row.put('fullName', getNodeText(rule, 'fullName'));
                row.put('active', getNodeText(rule, 'active'));
                row.put('triggerType', getNodeText(rule, 'triggerType'));
                row.put('description', getNodeText(rule, 'description'));
                row.put('formula', getNodeText(rule, 'formula'));

                List<String> criteria = parseCriteriaItems(rule);
                row.put('criteriaItems', String.join(criteria, '; '));

                results.add(row);
            }
            return results;
        }
    }

    public class ValidationRuleParser implements IMetadataParser {
        public List<Map<String, Object>> parse(String xmlContent, String filePath, 
                                               MetadataTypeConfig.TypeConfig config) {
            List<Map<String, Object>> results = new List<Map<String, Object>>();
            String objectName = extractObjectName(filePath);

            Dom.Document doc = new Dom.Document();
            doc.load(xmlContent);
            Dom.XmlNode root = doc.getRootElement();

            for (Dom.XmlNode rule : getChildNodes(root, 'validationRules')) {
                Map<String, Object> row = new Map<String, Object>();
                row.put('objectName', objectName);
                row.put('fullName', getNodeText(rule, 'fullName'));
                row.put('active', getNodeText(rule, 'active'));
                row.put('description', getNodeText(rule, 'description'));
                row.put('errorConditionFormula', getNodeText(rule, 'errorConditionFormula'));
                row.put('errorMessage', getNodeText(rule, 'errorMessage'));
                row.put('errorDisplayField', getNodeText(rule, 'errorDisplayField'));

                results.add(row);
            }
            return results;
        }
    }

    public class ApprovalProcessParser implements IMetadataParser {
        public List<Map<String, Object>> parse(String xmlContent, String filePath, 
                                               MetadataTypeConfig.TypeConfig config) {
            List<Map<String, Object>> results = new List<Map<String, Object>>();
            String objectName = extractObjectName(filePath);

            Dom.Document doc = new Dom.Document();
            doc.load(xmlContent);
            Dom.XmlNode root = doc.getRootElement();

            Map<String, Object> row = new Map<String, Object>();
            row.put('objectName', objectName);
            row.put('fullName', getNodeText(root, 'fullName'));
            row.put('label', getNodeText(root, 'label'));
            row.put('active', getNodeText(root, 'active'));
            row.put('recordEditability', getNodeText(root, 'recordEditability'));
            row.put('allowRecall', getNodeText(root, 'allowRecall'));

            results.add(row);
            return results;
        }
    }

    public class FlowParser implements IMetadataParser {
        public List<Map<String, Object>> parse(String xmlContent, String filePath, 
                                               MetadataTypeConfig.TypeConfig config) {
            List<Map<String, Object>> results = new List<Map<String, Object>>();
            String name = extractObjectName(filePath);

            Dom.Document doc = new Dom.Document();
            doc.load(xmlContent);
            Dom.XmlNode root = doc.getRootElement();

            Map<String, Object> row = new Map<String, Object>();
            row.put('fullName', name);
            row.put('label', getNodeText(root, 'label'));
            row.put('processType', getNodeText(root, 'processType'));
            row.put('status', getNodeText(root, 'status'));
            row.put('startElementReference', getNodeText(root, 'startElementReference'));
            row.put('apiVersion', getNodeText(root, 'apiVersion'));
            row.put('description', getNodeText(root, 'description'));

            results.add(row);
            return results;
        }
    }

    public class PermissionSetParser implements IMetadataParser {
        public List<Map<String, Object>> parse(String xmlContent, String filePath, 
                                               MetadataTypeConfig.TypeConfig config) {
            List<Map<String, Object>> results = new List<Map<String, Object>>();
            String name = extractObjectName(filePath);

            Dom.Document doc = new Dom.Document();
            doc.load(xmlContent);
            Dom.XmlNode root = doc.getRootElement();

            Map<String, Object> row = new Map<String, Object>();
            row.put('fullName', name);
            row.put('label', getNodeText(root, 'label'));
            row.put('description', getNodeText(root, 'description'));
            row.put('license', getNodeText(root, 'license'));
            row.put('fieldPermissionsCount', String.valueOf(getChildNodes(root, 'fieldPermissions').size()));
            row.put('objectPermissionsCount', String.valueOf(getChildNodes(root, 'objectPermissions').size()));

            results.add(row);
            return results;
        }
    }

    public class ProfileParser implements IMetadataParser {
        public List<Map<String, Object>> parse(String xmlContent, String filePath, 
                                               MetadataTypeConfig.TypeConfig config) {
            List<Map<String, Object>> results = new List<Map<String, Object>>();
            String name = extractObjectName(filePath);

            Dom.Document doc = new Dom.Document();
            doc.load(xmlContent);
            Dom.XmlNode root = doc.getRootElement();

            Map<String, Object> row = new Map<String, Object>();
            row.put('fullName', name);
            row.put('description', getNodeText(root, 'description'));
            row.put('userLicense', getNodeText(root, 'userLicense'));
            row.put('custom', getNodeText(root, 'custom'));
            row.put('fieldPermissionsCount', String.valueOf(getChildNodes(root, 'fieldPermissions').size()));
            row.put('objectPermissionsCount', String.valueOf(getChildNodes(root, 'objectPermissions').size()));

            results.add(row);
            return results;
        }
    }

    public class SharingRuleParser implements IMetadataParser {
        public List<Map<String, Object>> parse(String xmlContent, String filePath, 
                                               MetadataTypeConfig.TypeConfig config) {
            List<Map<String, Object>> results = new List<Map<String, Object>>();
            String objectName = extractObjectName(filePath);

            Dom.Document doc = new Dom.Document();
            doc.load(xmlContent);
            Dom.XmlNode root = doc.getRootElement();

            for (Dom.XmlNode rule : getChildNodes(root, 'sharingCriteriaRules')) {
                Map<String, Object> row = new Map<String, Object>();
                row.put('objectName', objectName);
                row.put('fullName', getNodeText(rule, 'fullName'));
                row.put('label', getNodeText(rule, 'label'));
                row.put('accessLevel', getNodeText(rule, 'accessLevel'));
                results.add(row);
            }

            for (Dom.XmlNode rule : getChildNodes(root, 'sharingOwnerRules')) {
                Map<String, Object> row = new Map<String, Object>();
                row.put('objectName', objectName);
                row.put('fullName', getNodeText(rule, 'fullName'));
                row.put('label', getNodeText(rule, 'label'));
                row.put('accessLevel', getNodeText(rule, 'accessLevel'));
                results.add(row);
            }

            return results;
        }
    }
}