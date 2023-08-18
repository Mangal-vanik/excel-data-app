// components/ExcelFileReader.tsx
"use client";
import React, { useState } from "react";
import { Workbook } from "exceljs";
import TextField from "@mui/material/TextField";
import { Box, Button, Typography } from "@mui/material";

import { useFormik } from "formik";
import * as Yup from "yup";
import {
  ICsvResult,
  getAccumulatedResults,
  getDataFromCsv,
  getDataFromExcel,
} from "./helper";

const validationSchema = Yup.object().shape({
  field1: Yup.mixed().required("Field 1 is required"),
  field2: Yup.mixed().required("Field 2 is required"),
});

function ExcelFileReader() {
  const [workbook, setWorkbook] = useState<Workbook>(new Workbook());

  const formik = useFormik({
    initialValues: {
      field1: "",
      field2: "",
    },
    validationSchema,
    onSubmit: async (values) => {
      const qty = getDataFromExcel(workbook);
      const csvResults = (await getDataFromCsv(values.field2)) as ICsvResult[];
      const finalResults = getAccumulatedResults(qty, csvResults, workbook);

      console.log("results", finalResults);
    },
  });

  const handleFileChange = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0] as File;

    if (file) {
      formik.setFieldValue("field1", file);
      workbook.xlsx.load(await file.arrayBuffer());
      setWorkbook(workbook);
    }
  };

  const handleChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0] as File;

    if (file) {
      formik.setFieldValue("field2", file);
    }
  };

  return (
    <Box
      style={{
        height: "inherit",
        display: "flex",
        flexDirection: "column",
        justifyContent: "center",
        alignItems: "center",
      }}
    >
      <Typography
        style={{
          padding: "30px",
        }}
        variant="h5"
      >
        EXCEL WEB APP
      </Typography>
      <Box
        component={"form"}
        onSubmit={formik.handleSubmit}
        style={{
          display: "flex",
          flexDirection: "column",
        }}
      >
        <label htmlFor="field1">Enter axlx file</label>

        <TextField
          label="Field 1:"
          variant="standard"
          type="file"
          onChange={handleFileChange}
          inputProps={{
            accept: ".xlsx",
          }}
          id="field1"
          name="field1"
          onBlur={formik.handleBlur}
          error={formik.touched.field1 && !!formik.errors.field1}
          helperText={formik.touched.field1 && formik.errors.field1}
        />
        <br />

        <div
          style={{
            marginTop: "5px",
          }}
        >
          <label htmlFor="field2">Enter csv file:</label>
          <br />

          <TextField
            onChange={handleChange}
            inputProps={{
              accept: ".csv",
            }}
            label="Field 2:"
            variant="standard"
            type="file"
            id="field2"
            name="field2"
            onBlur={formik.handleBlur}
            error={formik.touched.field2 && !!formik.errors.field2}
            helperText={formik.touched.field2 && formik.errors.field2}
          />
        </div>
        <Button
          variant="contained"
          type="submit"
          sx={{
            backgroundColor: "blue",
            color: "white",
            padding: "8px",
            marginTop: "20px",
            "&:hover": {
              backgroundColor: "darkblue",
            },
          }}
        >
          Submit
        </Button>
      </Box>
    </Box>
  );
}

export default ExcelFileReader;
