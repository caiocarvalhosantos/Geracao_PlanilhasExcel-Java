/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.arquivo_auxiliar_actionlistening;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class Java_Excel {
    public static void main(String[] args) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream arquivoSaida = new FileOutputStream("Relatorio_Alunos.xlsx")) {

            Sheet planilha = workbook.createSheet("Cadastro");
            
            Row cabecalho = planilha.createRow(0);
            cabecalho.createCell(0).setCellValue("RM");
            cabecalho.createCell(1).setCellValue("Nome");
            cabecalho.createCell(2).setCellValue("Curso");

            Row linhaDados = planilha.createRow(1);
            linhaDados.createCell(0).setCellValue(10010);
            linhaDados.createCell(1).setCellValue("João");
            linhaDados.createCell(2).setCellValue("Ciência da Computação");

            workbook.write(arquivoSaida);
            System.out.println("Planilha gerada com sucesso!");
        } catch (IOException e) {
            System.err.println("Erro ao salvar o arquivo: " + e.getMessage());
        }
    }
}
