package ru.tecon.uniRep;

import jakarta.ejb.EJB;
import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import ru.tecon.uniRep.ejb.UniRepBean;
import java.io.IOException;

@WebServlet("/loadUniRep")
public class Servlet extends HttpServlet {

    @EJB
    private UniRepBean uniRepBean;

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        int reportId = Integer.parseInt(req.getParameter("reportId"));

        uniRepBean.createReport(reportId);


        resp.setStatus(HttpServletResponse.SC_OK);
    }

}
