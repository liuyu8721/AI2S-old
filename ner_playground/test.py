from transformers import BertTokenizer, BertModel

model_path = "../models/chinese-pert-large/"

tokenizer = BertTokenizer.from_pretrained(model_path, from_tf=True)
model = BertModel.from_pretrained(model_path, from_tf=True)

inputs = tokenizer(r"""下午乘出租车到三亚市中心医院发热门诊就诊 住院隔离""", return_tensors="pt")
outputs = model(**inputs)

print(outputs)
