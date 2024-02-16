package ru.tecon.uniRep;

import ru.tecon.uniRep.ejb.UniRepBean;
//import ru.tecon.uniRep.ejb.UniRepSB;

import javax.ejb.EJB;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@WebServlet("/loadUniRep")
public class Servlet extends HttpServlet {

    @EJB
    private UniRepBean uniRepBean;
//    @EJB
//    private UniRepSB uniRepBean;

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        int reportId = Integer.parseInt(req.getParameter("reportId"));

        uniRepBean.createReport(reportId);


        resp.setStatus(HttpServletResponse.SC_OK);
    }

}
