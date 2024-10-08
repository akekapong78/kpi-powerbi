import requests
import os
from datetime import datetime
from requests_ntlm import HttpNtlmAuth

# NTLM authentication credentials
username = os.environ.get("USERNAME")
password = os.environ.get("MY_PASSWORD")
type = input("กรุณากรอก กฟฟ หรือ กฟส : ")
query = input("กรุณากรอกชื่อเต็ม นราธิวาส/รือเสาะ/ตากใบ/ระแงะ : ")
# query = "นราธิวาส"


# URL for the Power BI report
report_url = "https://peaex.pea.co.th/reports/powerbi/PRD/%E0%B8%81%E0%B8%9F%E0%B8%95.3%20%E0%B8%A2%E0%B8%B0%E0%B8%A5%E0%B8%B2/01.KPI%20%E0%B8%9C%E0%B8%88%E0%B8%81.%2067"

# URL for exporting the report in XLSX format (from your earlier example)
export_url = "https://peaex.pea.co.th/powerbi/api/explore/reports/a92f5b22-f71e-4c13-84fb-b69ea82f7b05/export/xlsx"
# Create a session to persist authentication
session = requests.Session()

# Perform NTLM authentication
session.auth = HttpNtlmAuth(username, password)

# Send request to authenticate and access the report page
response = session.get(report_url)

custom_body = ""
if type == "กฟส":
    custom_body = {
                "exportDataType": 0,
                "executeSemanticQueryRequest": {
                    "version": "1.0.0",
                    "queries": [
                        {
                            "Query": {
                                "Commands": [
                                    {
                                        "SemanticQueryDataShapeCommand": {
                                            "Query": {
                                                "Version": 2,
                                                "From": [
                                                    {
                                                        "Name": "d1",
                                                        "Entity": "DATA กฟส",
                                                        "Type": 0
                                                    },
                                                    {
                                                        "Name": "d11",
                                                        "Entity": "DATAxW",
                                                        "Type": 0
                                                    }
                                                ],
                                                "Select": [
                                                    {
                                                        "Aggregation": {
                                                            "Expression": {
                                                                "Column": {
                                                                    "Expression": {
                                                                        "SourceRef": {
                                                                            "Source": "d1"
                                                                        }
                                                                    },
                                                                    "Property": "ลำดับที่"
                                                                }
                                                            },
                                                            "Function": 0
                                                        },
                                                        "Name": "Sum(DATA กฟส.ลำดับที่)",
                                                        "NativeReferenceName": "ที่"
                                                    },
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d1"
                                                                }
                                                            },
                                                            "Property": "ข้อ"
                                                        },
                                                        "Name": "DATA กฟส.ข้อ",
                                                        "NativeReferenceName": "ข้อ"
                                                    },
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d1"
                                                                }
                                                            },
                                                            "Property": "การควบคุมภายใน"
                                                        },
                                                        "Name": "DATA กฟส.การควบคุมภายใน",
                                                        "NativeReferenceName": "การควบคุมภายใน"
                                                    },
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d1"
                                                                }
                                                            },
                                                            "Property": "ผู้รับผิดชอบ"
                                                        },
                                                        "Name": "DATA กฟส.ผู้รับผิดชอบ",
                                                        "NativeReferenceName": "ผู้รับผิดชอบ"
                                                    },
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d1"
                                                                }
                                                            },
                                                            "Property": "สถานะสะสม"
                                                        },
                                                        "Name": "DATA กฟส.สถานะสะสม",
                                                        "NativeReferenceName": "สถานะสะสม"
                                                    },
                                                    {
                                                        "Aggregation": {
                                                            "Expression": {
                                                                "Column": {
                                                                    "Expression": {
                                                                        "SourceRef": {
                                                                            "Source": "d1"
                                                                        }
                                                                    },
                                                                    "Property": "น้ำหนัก %"
                                                                }
                                                            },
                                                            "Function": 0
                                                        },
                                                        "Name": "Sum(DATA กฟส.น้ำหนัก %)",
                                                        "NativeReferenceName": "น้ำหนัก %1"
                                                    },
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d1"
                                                                }
                                                            },
                                                            "Property": "Link"
                                                        },
                                                        "Name": "DATA กฟส.Link",
                                                        "NativeReferenceName": "Link"
                                                    },
                                                    {
                                                        "Aggregation": {
                                                            "Expression": {
                                                                "Column": {
                                                                    "Expression": {
                                                                        "SourceRef": {
                                                                            "Source": "d1"
                                                                        }
                                                                    },
                                                                    "Property": "ระดับเกณฑ์ประเมิน"
                                                                }
                                                            },
                                                            "Function": 0
                                                        },
                                                        "Name": "Sum(DATA กฟส.ระดับเกณฑ์ประเมิน)",
                                                        "NativeReferenceName": "ระดับเกณฑ์ประเมิน1"
                                                    }
                                                ],
                                                "Where": [
                                                    {
                                                        "Condition": {
                                                            "In": {
                                                                "Expressions": [
                                                                    {
                                                                        "Column": {
                                                                            "Expression": {
                                                                                "SourceRef": {
                                                                                    "Source": "d1"
                                                                                }
                                                                            },
                                                                            "Property": "กฟฟ."
                                                                        }
                                                                    }
                                                                ],
                                                                "Values": [
                                                                    [
                                                                        {
                                                                            "Literal": {
                                                                                "Value": f"'{query}'"
                                                                            }
                                                                        }
                                                                    ]
                                                                ]
                                                            }
                                                        }
                                                    },
                                                    {
                                                        "Condition": {
                                                            "In": {
                                                                "Expressions": [
                                                                    {
                                                                        "Column": {
                                                                            "Expression": {
                                                                                "SourceRef": {
                                                                                    "Source": "d11"
                                                                                }
                                                                            },
                                                                            "Property": "ชั้น กฟฟ."
                                                                        }
                                                                    }
                                                                ],
                                                                "Values": [
                                                                    [
                                                                        {
                                                                            "Literal": {
                                                                                "Value": "'กฟส.'"
                                                                            }
                                                                        }
                                                                    ]
                                                                ]
                                                            }
                                                        }
                                                    }
                                                ],
                                                "OrderBy": [
                                                    {
                                                        "Direction": 1,
                                                        "Expression": {
                                                            "Aggregation": {
                                                                "Expression": {
                                                                    "Column": {
                                                                        "Expression": {
                                                                            "SourceRef": {
                                                                                "Source": "d1"
                                                                            }
                                                                        },
                                                                        "Property": "ลำดับที่"
                                                                    }
                                                                },
                                                                "Function": 0
                                                            }
                                                        }
                                                    }
                                                ]
                                            },
                                            "Binding": {
                                                "Primary": {
                                                    "Groupings": [
                                                        {
                                                            "Projections": [
                                                                0,
                                                                1,
                                                                2,
                                                                3,
                                                                4,
                                                                5,
                                                                6,
                                                                7
                                                            ],
                                                            "Subtotal": 0
                                                        }
                                                    ]
                                                },
                                                "DataReduction": {
                                                    "Primary": {
                                                        "Top": {
                                                            "Count": 1000000
                                                        }
                                                    },
                                                    "Secondary": {
                                                        "Top": {
                                                            "Count": 100
                                                        }
                                                    }
                                                },
                                                "Version": 1
                                            }
                                        }
                                    },
                                    {
                                        "ExportDataCommand": {
                                            "Columns": [
                                                {
                                                    "QueryName": "Sum(DATA กฟส.ลำดับที่)",
                                                    "Name": "ที่"
                                                },
                                                {
                                                    "QueryName": "DATA กฟส.ข้อ",
                                                    "Name": "ข้อ"
                                                },
                                                {
                                                    "QueryName": "DATA กฟส.การควบคุมภายใน",
                                                    "Name": "การควบคุมภายใน"
                                                },
                                                {
                                                    "QueryName": "DATA กฟส.ผู้รับผิดชอบ",
                                                    "Name": "ผู้รับผิดชอบ"
                                                },
                                                {
                                                    "QueryName": "DATA กฟส.สถานะสะสม",
                                                    "Name": "สถานะสะสม"
                                                },
                                                {
                                                    "QueryName": "Sum(DATA กฟส.น้ำหนัก %)",
                                                    "Name": "น้ำหนัก %"
                                                },
                                                {
                                                    "QueryName": "DATA กฟส.Link",
                                                    "Name": "Link"
                                                },
                                                {
                                                    "QueryName": "Sum(DATA กฟส.ระดับเกณฑ์ประเมิน)",
                                                    "Name": "ระดับเกณฑ์ประเมิน"
                                                }
                                            ],
                                            "Ordering": [
                                                0,
                                                1,
                                                2,
                                                3,
                                                4,
                                                5,
                                                6,
                                                7
                                            ],
                                            "FiltersDescription": f"Applied filters:\nกฟฟ. is {query}"
                                        }
                                    }
                                ]
                            }
                        }
                    ],
                    "cancelQueries": [],
                    "modelId": "1227254892",
                    "userPreferredLocale": "en-US"
                }
            }

if type == "กฟฟ":
    custom_body = {
                "exportDataType": 0,
                "executeSemanticQueryRequest": {
                    "version": "1.0.0",
                    "queries": [
                        {
                            "Query": {
                                "Commands": [
                                    {
                                        "SemanticQueryDataShapeCommand": {
                                            "Query": {
                                                "Version": 2,
                                                "From": [
                                                    {
                                                        "Name": "d",
                                                        "Entity": "DATA กฟฟ",
                                                        "Type": 0
                                                    },
                                                    {
                                                        "Name": "d1",
                                                        "Entity": "DATAxW",
                                                        "Type": 0
                                                    }
                                                ],
                                                "Select": [
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d"
                                                                }
                                                            },
                                                            "Property": "ข้อ"
                                                        },
                                                        "Name": "DATA กฟฟ.ข้อ",
                                                        "NativeReferenceName": "ข้อ"
                                                    },
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d"
                                                                }
                                                            },
                                                            "Property": "การควบคุมภายใน"
                                                        },
                                                        "Name": "DATA กฟฟ.การควบคุมภายใน",
                                                        "NativeReferenceName": "การควบคุมภายใน"
                                                    },
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d"
                                                                }
                                                            },
                                                            "Property": "ผู้รับผิดชอบ"
                                                        },
                                                        "Name": "DATA กฟฟ.ผู้รับผิดชอบ",
                                                        "NativeReferenceName": "ผู้รับผิดชอบ"
                                                    },
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d"
                                                                }
                                                            },
                                                            "Property": "สถานะสะสม"
                                                        },
                                                        "Name": "DATA กฟฟ.สถานะสะสม",
                                                        "NativeReferenceName": "สถานะสะสม"
                                                    },
                                                    {
                                                        "Aggregation": {
                                                            "Expression": {
                                                                "Column": {
                                                                    "Expression": {
                                                                        "SourceRef": {
                                                                            "Source": "d"
                                                                        }
                                                                    },
                                                                    "Property": "น้ำหนัก %"
                                                                }
                                                            },
                                                            "Function": 0
                                                        },
                                                        "Name": "Sum(DATA กฟฟ.น้ำหนัก %)",
                                                        "NativeReferenceName": "น้ำหนัก %"
                                                    },
                                                    {
                                                        "Column": {
                                                            "Expression": {
                                                                "SourceRef": {
                                                                    "Source": "d"
                                                                }
                                                            },
                                                            "Property": "Link"
                                                        },
                                                        "Name": "DATA กฟฟ.Link",
                                                        "NativeReferenceName": "Link"
                                                    },
                                                    {
                                                        "Aggregation": {
                                                            "Expression": {
                                                                "Column": {
                                                                    "Expression": {
                                                                        "SourceRef": {
                                                                            "Source": "d"
                                                                        }
                                                                    },
                                                                    "Property": "ระดับเกณฑ์ประเมิน"
                                                                }
                                                            },
                                                            "Function": 0
                                                        },
                                                        "Name": "Sum(DATA กฟฟ.ระดับเกณฑ์ประเมิน)",
                                                        "NativeReferenceName": "ระดับเกณฑ์ประเมิน"
                                                    },
                                                    {
                                                        "Aggregation": {
                                                            "Expression": {
                                                                "Column": {
                                                                    "Expression": {
                                                                        "SourceRef": {
                                                                            "Source": "d"
                                                                        }
                                                                    },
                                                                    "Property": "ลำดับที่"
                                                                }
                                                            },
                                                            "Function": 0
                                                        },
                                                        "Name": "Sum(DATA กฟฟ.ลำดับที่)",
                                                        "NativeReferenceName": "ที่"
                                                    }
                                                ],
                                                "Where": [
                                                    {
                                                        "Condition": {
                                                            "In": {
                                                                "Expressions": [
                                                                    {
                                                                        "Column": {
                                                                            "Expression": {
                                                                                "SourceRef": {
                                                                                    "Source": "d"
                                                                                }
                                                                            },
                                                                            "Property": "กฟฟ."
                                                                        }
                                                                    }
                                                                ],
                                                                "Values": [
                                                                    [
                                                                        {
                                                                            "Literal": {
                                                                                "Value": f"'{query}'"
                                                                            }
                                                                        }
                                                                    ]
                                                                ]
                                                            }
                                                        }
                                                    },
                                                    {
                                                        "Condition": {
                                                            "In": {
                                                                "Expressions": [
                                                                    {
                                                                        "Column": {
                                                                            "Expression": {
                                                                                "SourceRef": {
                                                                                    "Source": "d1"
                                                                                }
                                                                            },
                                                                            "Property": "ชั้น กฟฟ."
                                                                        }
                                                                    }
                                                                ],
                                                                "Values": [
                                                                    [
                                                                        {
                                                                            "Literal": {
                                                                                "Value": "'กฟฟ.'"
                                                                            }
                                                                        }
                                                                    ]
                                                                ]
                                                            }
                                                        }
                                                    }
                                                ],
                                                "OrderBy": [
                                                    {
                                                        "Direction": 1,
                                                        "Expression": {
                                                            "Aggregation": {
                                                                "Expression": {
                                                                    "Column": {
                                                                        "Expression": {
                                                                            "SourceRef": {
                                                                                "Source": "d"
                                                                            }
                                                                        },
                                                                        "Property": "ลำดับที่"
                                                                    }
                                                                },
                                                                "Function": 0
                                                            }
                                                        }
                                                    }
                                                ]
                                            },
                                            "Binding": {
                                                "Primary": {
                                                    "Groupings": [
                                                        {
                                                            "Projections": [
                                                                0,
                                                                1,
                                                                2,
                                                                3,
                                                                4,
                                                                5,
                                                                6,
                                                                7
                                                            ],
                                                            "Subtotal": 0
                                                        }
                                                    ]
                                                },
                                                "DataReduction": {
                                                    "Primary": {
                                                        "Top": {
                                                            "Count": 1000000
                                                        }
                                                    },
                                                    "Secondary": {
                                                        "Top": {
                                                            "Count": 100
                                                        }
                                                    }
                                                },
                                                "Version": 1
                                            }
                                        }
                                    },
                                    {
                                        "ExportDataCommand": {
                                            "Columns": [
                                                {
                                                    "QueryName": "DATA กฟฟ.ข้อ",
                                                    "Name": "ข้อ"
                                                },
                                                {
                                                    "QueryName": "DATA กฟฟ.การควบคุมภายใน",
                                                    "Name": "การควบคุมภายใน"
                                                },
                                                {
                                                    "QueryName": "DATA กฟฟ.ผู้รับผิดชอบ",
                                                    "Name": "ผู้รับผิดชอบ"
                                                },
                                                {
                                                    "QueryName": "DATA กฟฟ.สถานะสะสม",
                                                    "Name": "สถานะสะสม"
                                                },
                                                {
                                                    "QueryName": "Sum(DATA กฟฟ.น้ำหนัก %)",
                                                    "Name": "น้ำหนัก %"
                                                },
                                                {
                                                    "QueryName": "DATA กฟฟ.Link",
                                                    "Name": "Link"
                                                },
                                                {
                                                    "QueryName": "Sum(DATA กฟฟ.ระดับเกณฑ์ประเมิน)",
                                                    "Name": "ระดับเกณฑ์ประเมิน"
                                                },
                                                {
                                                    "QueryName": "Sum(DATA กฟฟ.ลำดับที่)",
                                                    "Name": "ที่"
                                                }
                                            ],
                                            "Ordering": [
                                                7,
                                                0,
                                                1,
                                                2,
                                                3,
                                                4,
                                                5,
                                                6
                                            ],
                                            "FiltersDescription": f"Applied filters:\nกฟฟ. is {query}"
                                        }
                                    }
                                ]
                            }
                        }
                    ],
                    "cancelQueries": [],
                    "modelId": "1227254892",
                    "userPreferredLocale": "en-US"
                }
            }

# Check if authentication was successful
if response.status_code == 200:
    print("Login successful. Accessing the report...")

    # Now, request to export the report
    export_response = session.post(
        export_url,
        headers={
            "accept": "application/json, text/plain, */*",
            "content-type": "application/json;charset=UTF-8"
        },
        json=custom_body
    )

    # Check if the report export request was successful
    if export_response.status_code == 200:
        print("Report export successful. Saving the file...")

        folder_name = datetime.now().strftime("%m.%Y")
        current_date = datetime.now().strftime("%d.%m.%Y")

        # Create the folder
        os.makedirs(folder_name, exist_ok=True)

        # Save the XLSX report to a file
        report_name = f"report_{query}_{current_date}.xlsx"
        save_path = os.path.join(folder_name, report_name)

        with open(save_path, "wb") as file:
            file.write(export_response.content)

        print(f"Report saved as '{report_name}'.")
    else:
        print(f"Failed to export report. Status code: {export_response.status_code}")
else:
    print(f"Failed to login. Status code: {response.status_code}")
