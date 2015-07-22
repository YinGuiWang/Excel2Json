根据目标配置一个结构，格式如Config.json 所示

格式如下：（默认每个表需有一个主键，且在第一列）
[
	
    {
        "excel_path": "D:/CAT/cat_server/work_help_script/excel2Json/Skill.xlsx",  #excel路径
        "main_sheet": "Skill",  #excel主sheet名称
        "output_path": "D:/CAT/cat_server/work_help_script/excel2Json/Skill", #输出json 文件位置
        "is_merge":1 #是否合并 0,不合并 1，合并
    },
    {
        "id": {					#普通key-value 配表方法
            "sheet": "Skill",   #所在sheet名称
            "valueCol": 0,      #值在第几列
            "result": "value"   #与取得的值处理
        },
        "inDisMin": {
            "sheet": "Skill",
            "valueCol": 9,
            "result": "value / 10000"
        },
        "inDisMax": {
            "sheet": "Skill",
            "valueCol": 9,
            "result": "value / 10000"
        },
        "cd": {
            "sheet": "Skill",
            "valueCol": 4,
            "result": "value"
        },
        "phases": [										#普通List 配表方法
            {                                           
                "type": 1,                              #list 类型   1:在表格中有多行的  2:在表格中是一行的 
                "sheet": "skillData",                   #所在sheet名称
                "commonCol": 1                          #公共key值在第几列
            },
            {
                "duration": {							
                    "sheet": "skillData",
                    "valueCol": 4,
                    "result": "value"
                },
                "collision": {
                    "hitId": {
                        "sheet": "skillData",
                        "valueCol": 2,
                        "result": "value"
                    },
                    "hitType": {
                        "sheet": "skillData",
                        "valueCol": 24,
                        "result": "value"
                    },
                    "controlType": {
                        "sheet": "skillData",
                        "valueCol": 23,
                        "result": "value"
                    },
                    "valIdx": {
                        "sheet": "skillData",
                        "valueCol": 2,
                        "result": "value - 1"
                    },
                    "deltaTime": {
                        "sheet": "skillData",
                        "valueCol": 4,
                        "result": "value"
                    },
                    "r": {
                        "sheet": "skillData",
                        "valueCol": 26,
                        "result": "value"
                    },
                    "circles": [
                        {
                            "type": 2,
                            "sheet": "skillData",
                            "commonCol": 28,
                            "keys": [
                                "x",
                                "y",
                                "z"
                            ]
                        }
                    ]
                }
            }
        ],
        "vals": [
            {
                "type": 1,
                "sheet": "skillData",
                "commonCol": 1
            },
            {
                "atk": {
                    "sheet": "skillData",
                    "valueCol": 9,
                    "result": "value"
                },
                "atkPer": {
                    "sheet": "skillData",
                    "valueCol": 8,
                    "result": "value / 100"
                },
                "superArmorDmg": {
                    "sheet": "skillData",
                    "valueCol": 10,
                    "result": "value"
                },
                "superArmorIgnore": {
                    "sheet": "skillData",
                    "valueCol": 7,
                    "result": "value"
                },
                "control": {
                    "sheet": "skillData",
                    "valueCol": 41,
                    "result": "value"
                },
                "controlDuration": {
                    "sheet": "skillData",
                    "valueCol": 45,
                    "other[0]": 46,			#有多个值的 除第一个外，其他顺延 other[0] other[1] other[2]......
                    "result": "max(value, other[0])"
                },
                "buffs": {
                    "sheet": "skillData",
                    "valueCol": 13,
                    "result": "value"
                },
                "dots": {
                    "sheet": "skillData",
                    "valueCol": 14,
                    "result": "value"
                }
            }
        ]
    }
]