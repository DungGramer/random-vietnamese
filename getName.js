const XLSX = require("xlsx");
const fs = require("fs");

// const workbook = XLSX.readFile("pl_ii_3112202116.xlsx");
// const workbook = XLSX.readFile("Thong tin bang TN THPT 2018 - Cong bo.xlsx");

const sheet_name_list = workbook.SheetNames;
const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

const NAME_ROW = "Họ và tên";
const GENDER_ROW = "Giới tính";

const upperFirstLetter = (name) => {
  const splitWord = name.split(" ");

  if (splitWord.length === 1) {
    return name.charAt(0).toUpperCase() + name.slice(1).toLowerCase();
  }

  let result = "";
  for (const word of splitWord) {
    result += `${word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()} `;
  }
  return result;
};

const convertGender = (gender) => {
  switch (gender) {
    case "Nam":
      return "M";
    case "Nữ":
      return "F";
    default:
      return "O";
  }
};

const getFirstName = (name) => {
  return upperFirstLetter(name.split(" ")[0]);
};

const getLastName = (name) => {
  return upperFirstLetter(name.split(" ").pop());
};

const getMiddleName = (name) => {
  const middleName = name.split(" ");
  middleName.shift();
  middleName.pop();
  return upperFirstLetter(middleName.join(" ").replace(/ +(?= )/g, ""));
};

const Name = JSON.parse(fs.readFileSync("name.json", "utf8")) || {
  firstName: [],
  lastName: [],
  middleName: [],
};
let FirstName = new Set();
let LastName = new Set();
let MiddleName = new Set();
const Gender = [];

xlData.map((obj) => {
  const name = obj[NAME_ROW];
  const gender = obj[GENDER_ROW];
  if (typeof name === "string" && !/\n|\t|\r|\f|\v/g.test(name)) {
    const firstName = getFirstName(name).trim();
    const lastName = getLastName(name).trim();
    const middleName = getMiddleName(name).trim();

    if (firstName && lastName && middleName && gender) {
      FirstName.add(firstName.trim());
      LastName.add(lastName.trim());
      MiddleName.add(middleName.trim());
      Gender.push(convertGender(gender));
    }
  }
});
FirstName = [...FirstName];
LastName = [...LastName];
MiddleName = [...MiddleName];

const FirstNameWithGender = [];
const LastNameWithGender = [];
const MiddleNameWithGender = [];

for (let i = 0; i < FirstName.length; i++) {
  FirstNameWithGender.push({
    name: FirstName[i],
    gender: Gender[i],
  });

  LastNameWithGender.push({
    name: LastName[i],
    gender: Gender[i],
  });

  MiddleNameWithGender.push({
    name: MiddleName[i],
    gender: Gender[i],
  });
}

const newFirstName = new Set(
  Name.firstName.concat(Array.from(FirstNameWithGender))
);
const newLastName = new Set(
  Name.lastName.concat(Array.from(LastNameWithGender))
);
const newMiddleName = new Set(
  Name.middleName.concat(Array.from(MiddleNameWithGender))
);

Name.firstName = Array.from(newFirstName);
Name.lastName = Array.from(newLastName);
Name.middleName = Array.from(newMiddleName);

fs.writeFileSync("name.json", JSON.stringify(Name), {
  encoding: "utf-8",
  flag: "w",
});
