package com.sixdglobal.aem.reports.servlets;

import org.apache.felix.scr.annotations.*;
import org.apache.felix.scr.annotations.Property;
import org.apache.jackrabbit.api.JackrabbitSession;
import org.apache.jackrabbit.api.security.user.Authorizable;
import org.apache.jackrabbit.api.security.user.Group;
import org.apache.jackrabbit.api.security.user.UserManager;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.resource.ResourceResolver;
import org.apache.sling.api.resource.ResourceResolverFactory;
import org.apache.sling.api.servlets.SlingAllMethodsServlet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import javax.jcr.*;
import javax.jcr.query.Query;
import javax.jcr.query.QueryManager;
import javax.jcr.query.QueryResult;
import javax.servlet.Servlet;
import javax.servlet.ServletException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.sling.jcr.api.SlingRepository;
import org.apache.commons.io.IOUtils;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

@Component(label="Group Enrollment Report", description="Parses users and their group enrollment then saves XLSX file", metatype = true, immediate = true)
@Service(value = { Servlet.class })
@Properties({
        @Property(name="sling.servlet.paths", value={"/bin/sixdglobal/groupenrollment"}, propertyPrivate=true),
        @Property(name="sling.servlet.methods", value={"POST"}, propertyPrivate=true)
})

public class GroupEnrollment extends SlingAllMethodsServlet {

    @Reference
    private ResourceResolverFactory resolverFactory;

    @Reference
    private SlingRepository repository;

    private static final Logger log = LoggerFactory.getLogger(GroupEnrollment.class);

    private String fileName = "group-report.xlsx";

    /**
     * doPost respond to POST requests.
     * Builds XLSX workbook and requests JCR save to binary node.
     *
     * @param request HTTP request
     * @param response HTTP response
     * @throws ServletException
     * @throws IOException
     */
    @Override
    protected void doPost(SlingHttpServletRequest request, SlingHttpServletResponse response) throws ServletException, IOException {
        //Request workbook object back from user query
        Workbook workbook = getUsers();

        //Make save attempt at the provided save path; return success or failure
        String savePath = null;
        if (request.getParameter("savePath") != null) {
            savePath = request.getParameter("savePath");
        }
        if (savePath!=null) {
            saveWorkbook(savePath, repository, workbook);
            response.getWriter().write("Success:\nFile being processed and saved.");
        } else {
            response.getWriter().write("Failure:\nMissing file save path.");
        }
    }

    /**
     * getUsers build and return Workbook object.
     * Query user data in JCR and build workbook of user's and their group enrollment.
     *
     * @return Workbook
     */
    public Workbook getUsers() {
        Session session = null;
        Workbook workbook = null;
        int row = 0;

        try
        {
            //Create workbook instance for XLSX file
            workbook = new SXSSFWorkbook(100);
            Sheet sheet = workbook.createSheet("Group Enrollment");

            //Build query and store results into rows
            ResourceResolver resourceResolver = resolverFactory.getAdministrativeResourceResolver(null);
            session = resourceResolver.adaptTo(Session.class);
            QueryManager queryManager = session.getWorkspace().getQueryManager();
            StringBuilder queryBuilder = new StringBuilder();
            queryBuilder.append("/jcr:root/home/users//element(*,rep:User)");
            Query query = queryManager.createQuery(queryBuilder.toString(), "xpath");
            QueryResult result = query.execute();
            NodeIterator nodeIter = result.getNodes();

            //Maintain array of groups per user to avoid reporting duplicates
            ArrayList groupList = new ArrayList();

            //For each user create a new row
            while ( nodeIter.hasNext() ) {
                int cell = 0;
                Row dataRow = sheet.createRow(row);
                Node node = nodeIter.nextNode();
                String userName = node.getName();
                Cell dataCell = dataRow.createCell(cell);
                dataCell.setCellValue(userName);

                groupList.clear();

                //For each user's group create a cell within the current row
                String groups = "";
                UserManager userManager = ((JackrabbitSession)session).getUserManager();
                Authorizable user = userManager.getAuthorizableByPath(node.getPath());
                Iterator<Group> groupIterator = user.memberOf();

                while (groupIterator.hasNext()) {
                    cell++;
                    dataCell = dataRow.createCell(cell);
                    Group group = groupIterator.next();
                    String groupName = group.getID();
                    groupList.add(groupName);
                    dataCell.setCellValue(groupName);

                    //Step through all transitive groups and create a cell within the current row
                    Iterator<Group> transitiveIterator1 = group.memberOf();
                    while (transitiveIterator1.hasNext()) {
                        group = transitiveIterator1.next();
                        groupName = group.getID();

                        //Check if group array contains this group; if not, add to the report
                        if (!groupList.contains(groupName)) {
                            cell++;
                            dataCell = dataRow.createCell(cell);
                            groupList.add(groupName);
                            dataCell.setCellValue(groupName);
                        }

                        //Find all transitive groups 2nd level
                        Iterator<Group> transitiveIterator2 = group.memberOf();
                        while (transitiveIterator2.hasNext()) {
                            group = transitiveIterator2.next();
                            groupName = group.getID();

                            //Check if group array contains this group; if not, add to the report
                            if (!groupList.contains(groupName)) {
                                cell++;
                                dataCell = dataRow.createCell(cell);
                                groupList.add(groupName);
                                dataCell.setCellValue(groupName);
                            }

                            //Find all transitive groups 3rd level
                            Iterator<Group> transitiveIterator3 = group.memberOf();
                            while (transitiveIterator3.hasNext()) {
                                group = transitiveIterator3.next();
                                groupName = group.getID();

                                //Check if group array contains this group; if not, add to the report
                                if (!groupList.contains(groupName)) {
                                    cell++;
                                    dataCell = dataRow.createCell(cell);
                                    groupList.add(groupName);
                                    dataCell.setCellValue(groupName);
                                }

                                //Find all transitive groups 4th level
                                Iterator<Group> transitiveIterator4 = group.memberOf();
                                while (transitiveIterator4.hasNext()) {
                                    group = transitiveIterator4.next();
                                    groupName = group.getID();

                                    //Check if group array contains this group; if not, add to the report
                                    if (!groupList.contains(groupName)) {
                                        cell++;
                                        dataCell = dataRow.createCell(cell);
                                        groupList.add(groupName);
                                        dataCell.setCellValue(groupName);
                                    }

                                    //Find all transitive groups 5th level
                                    Iterator<Group> transitiveIterator5 = group.memberOf();
                                    while (transitiveIterator5.hasNext()) {
                                        group = transitiveIterator5.next();
                                        groupName = group.getID();

                                        //Check if group array contains this group; if not, add to the report
                                        if (!groupList.contains(groupName)) {
                                            cell++;
                                            dataCell = dataRow.createCell(cell);
                                            groupList.add(groupName);
                                            dataCell.setCellValue(groupName);
                                        }

                                        //Find all transitive groups 6th level
                                        Iterator<Group> transitiveIterator6 = group.memberOf();
                                        while (transitiveIterator6.hasNext()) {
                                            group = transitiveIterator6.next();
                                            groupName = group.getID();

                                            //Check if group array contains this group; if not, add to the report
                                            if (!groupList.contains(groupName)) {
                                                cell++;
                                                dataCell = dataRow.createCell(cell);
                                                groupList.add(groupName);
                                                dataCell.setCellValue(groupName);
                                            }

                                            //Find all transitive groups 7th level
                                            Iterator<Group> transitiveIterator7 = group.memberOf();
                                            while (transitiveIterator7.hasNext()) {
                                                group = transitiveIterator7.next();
                                                groupName = group.getID();

                                                //Check if group array contains this group; if not, add to the report
                                                if (!groupList.contains(groupName)) {
                                                    cell++;
                                                    dataCell = dataRow.createCell(cell);
                                                    groupList.add(groupName);
                                                    dataCell.setCellValue(groupName);
                                                }

                                                //Find all transitive groups 8th level
                                                Iterator<Group> transitiveIterator8 = group.memberOf();
                                                while (transitiveIterator8.hasNext()) {
                                                    group = transitiveIterator8.next();
                                                    groupName = group.getID();

                                                    //Check if group array contains this group; if not, add to the report
                                                    if (!groupList.contains(groupName)) {
                                                        cell++;
                                                        dataCell = dataRow.createCell(cell);
                                                        groupList.add(groupName);
                                                        dataCell.setCellValue(groupName);
                                                    }

                                                    //Find all transitive groups 9th level
                                                    Iterator<Group> transitiveIterator9 = group.memberOf();
                                                    while (transitiveIterator9.hasNext()) {
                                                        group = transitiveIterator9.next();
                                                        groupName = group.getID();

                                                        //Check if group array contains this group; if not, add to the report
                                                        if (!groupList.contains(groupName)) {
                                                            cell++;
                                                            dataCell = dataRow.createCell(cell);
                                                            groupList.add(groupName);
                                                            dataCell.setCellValue(groupName);
                                                        }

                                                        //Find all transitive groups 10th level
                                                        Iterator<Group> transitiveIterator10 = group.memberOf();
                                                        while (transitiveIterator10.hasNext()) {
                                                            group = transitiveIterator10.next();
                                                            groupName = group.getID();

                                                            //Check if group array contains this group; if not, add to the report
                                                            if (!groupList.contains(groupName)) {
                                                                cell++;
                                                                dataCell = dataRow.createCell(cell);
                                                                groupList.add(groupName);
                                                                dataCell.setCellValue(groupName);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                row++;
            }
        } catch (RepositoryException e) {
            log.error("RepositoryException occurred while obtaining groups: "+e.getMessage());
        } catch(Exception e) {
            log.error("Exception occurred on GroupEnrollment.java: "+e.getMessage());
        }
        return workbook;
    }

    /**
     * saveWorkbook saves XLSX in desired JCR location
     * Stores the created Workbook object into a JCR binary node; force-overwrite of any existing node
     *
     * @param savePath Request parameter passed for desired save path within the JCR.
     * @param repository Repository object used to log admin session
     * @param workbook Create workbook object to save in the output stream.
     */
    protected void saveWorkbook(String savePath, SlingRepository repository, Workbook workbook){
        Session session = null;
        ByteArrayOutputStream outputStream = null;
        ByteArrayInputStream inputStream = null;

        try{
            session = repository.loginAdministrative(null);

            if(workbook != null){
                // write the generated file in memory so that it can be saved in the JCR
                outputStream = new ByteArrayOutputStream();
                workbook.write(outputStream);
                outputStream.flush();
                inputStream = new ByteArrayInputStream(outputStream.toByteArray());

                Node saveNode = session.getNode(savePath);
                if (saveNode.hasNode(fileName)) {
                    Node oldNode = saveNode.getNode(fileName);
                    oldNode.remove();
                    session.save();
                }
                Node fileNode = saveNode.addNode(fileName, "nt:file");
                fileNode.addMixin("mix:referenceable");
                Node contentNode = fileNode.addNode("jcr:content", "nt:resource");
                Binary binary = session.getValueFactory().createBinary(inputStream);
                contentNode.setProperty("jcr:mimeType",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                contentNode.setProperty("jcr:data", binary);
                inputStream.close();
                outputStream.close();
            }
            session.save();
        }catch(Exception e){
            log.error("exception: ", e);
        }finally {
            if(session != null){
                session.logout();
            }
            IOUtils.closeQuietly(inputStream);
            IOUtils.closeQuietly(outputStream);
        }
    }
}