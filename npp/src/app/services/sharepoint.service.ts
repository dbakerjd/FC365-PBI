import { NumberSymbol } from '@angular/common';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { AccountInfo } from '@azure/msal-browser';
import { ErrorService } from './error.service';
import { LicensingService } from './licensing.service';
import { TeamsService } from './teams.service';

export interface Opportunity {
  title: string;
  moleculeName: string;
  opportunityOwner: User;
  projectStart: Date;
  projectEnd: Date;
  opportunityType: string;
  opportunityStatus: string;
  indicationName: string;
  Id: number;
  therapyArea: string;
  updated: Date;
  users?: User[];
  progress: number;
}

export interface User {
  id: number;
  name: string;
  email?: string;
  profilePic?: string;
}

export interface Action {
  id: number,
  gateId: number;
  opportunityId: number;
  title: string;
  actionName: string;
  dueDate: Date;
  completed: boolean;
  timestamp: Date;
  targetUserId: Number;
  targetUser: User;
  status?: string;
}

export interface Gate {
  id: number;
  title: string;
  opportunityId: number;
  name: string;
  reviewedAt: Date;
  createdAt: Date;
  actions: Action[];
  folders?: NPPFolder[];
}

export interface NPPFile {
  id: number;
  parentId: number;
  name: string;
  updatedAt: Date;
  description: string;
  stageId: number;
  opportunityId: number;
  country: string[];
  modelScenario: string[];
  modelApprovalComments: string;
  approvalStatus: string;
  user: User;
}

export interface NPPFolder {
  id: number;
  name: string;
  containsModels?: boolean;
}

@Injectable({
  providedIn: 'root'
})
export class SharepointService {

  folders: NPPFolder[] = [{
    id: 1,
    name: 'Finance'
  },{
    id: 2,
    name: 'Commercial'
  }, {
    id: 3,
    name: 'Technical'
  }, {
    id: 4,
    name: 'Regulatory'
  }, {
    id: 5,
    name: 'Other'
  },{
    id: 6,
    name: 'Forecast Models',
    containsModels: true
  }];

  files: NPPFile[] = [{
    id: 1,
    parentId: 1,
    name: 'test.pdf',
    updatedAt: new Date(),
    description: 'test description',
    stageId: 1,
    opportunityId: 1,
    country: [],
    modelScenario: [],
    modelApprovalComments: '',
    approvalStatus: '',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  },{
    id: 2,
    parentId: 1,
    name: 'test2.pdf',
    updatedAt: new Date(),
    description: 'Another test description',
    stageId: 1,
    opportunityId: 1,
    country: [],
    modelScenario: [],
    modelApprovalComments: '',
    approvalStatus: '',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  },{
    id: 3,
    parentId: 1,
    name: 'test3.pdf',
    updatedAt: new Date(),
    description: 'Yet another test description',
    stageId: 1,
    opportunityId: 1,
    country: [],
    modelScenario: [],
    modelApprovalComments: '',
    approvalStatus: '',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  },{
    id: 4,
    parentId: 6,
    name: 'test_model',
    updatedAt: new Date(),
    description: 'Yet another test description',
    stageId: 1,
    opportunityId: 1,
    country: ['UK', 'Spain', 'Belgium'],
    modelScenario: ['Upside', 'Downside'],
    modelApprovalComments: 'Lorem Ipsum Dolor amet and all that',
    approvalStatus: 'In Progress',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  },{
    id: 5,
    parentId: 6,
    name: 'test_model3',
    updatedAt: new Date(),
    description: 'Yet another test description',
    stageId: 1,
    opportunityId: 1,
    country: ['UK', 'Spain', 'Belgium'],
    modelScenario: ['Upside', 'Downside'],
    modelApprovalComments: 'Some test random comment',
    approvalStatus: 'In Progress',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  }];

  opportunities: Opportunity[] =  [{
    title: "Acquisition of Nucala for COPD",
    moleculeName: "Nucala",
    opportunityOwner: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    },
    projectStart: new Date("5/1/2021"),
    projectEnd: new Date("11/1/2021"),
    opportunityType: "Acquisition",
    opportunityStatus: "Active",
    indicationName: "Chronic Obstructive Pulmonary Disease (COPD)",
    Id: 67,
    therapyArea: "Respiratory",
    updated: new Date("5/25/2021 3:04 PM"),
    progress: 79
  },{
    title: "Acquisition of Tezepelumab (Asthma)",
    moleculeName: "Tezepelumab",
    opportunityOwner: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    },
    projectStart: new Date("5/1/2021"),
    projectEnd: new Date("1/1/2024"),
    opportunityType: "Acquisition",
    opportunityStatus: "Active",
    indicationName: "Asthma",
    Id: 68,
    therapyArea: "Respiratory",
    updated: new Date("5/25/2021 3:55 PM"),
    progress: 45
  },{
    title: "Development of Concizumab",
    moleculeName: "Concizumab",
    opportunityOwner: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    },
    projectStart: new Date("5/1/2021"),
    projectEnd: new Date("5/1/2025"),
    opportunityType: "Product Development",
    opportunityStatus: "Archived",
    indicationName: "Hemophilia",
    Id: 69,
    therapyArea: "Haematology",
    updated: new Date("5/25/2021 4:14 PM"),
    progress: 12
  }];

  indications: any[]  = [
    { value: "1", label : "Bone Cancer" },
    { value: "2", label : "Bone Density" },
    { value: "3", label : "Bone Infections" },
    { value: "4", label : "Osteogenesis Imperfecta" },
    { value: "5", label : "Osteonecrosis" },
    { value: "6", label : "Osteoporosis"},
    { value: "7", label : "Paget's Disease"},
    { value: "8", label : "Rickets"},
    { value: "9", label : "Other"},
    { value: "10", label : "Arrhythmias"},
    { value: "11", label : "Aorta disease"},
    { value: "12", label : "Congenital heart disease"},
    { value: "13", label : "Coronary artery disease"},
    { value: "14", label : "DVT/ PE"},
    { value: "15", label : "Heart attack"},
    { value: "16", label : "Heart failure"},
    { value: "17", label : "Heart muscle disease"},
    { value: "18", label : "Heart valve disease"},
    { value: "19", label : "Pericardial disease"},
    { value: "20", label : "Peripheral vascular disease"},
    { value: "21", label : "Rheumatic heart disease"},
    { value: "22", label : "Stroke"},
    { value: "23", label : "Vascular disease"},
    { value: "24", label : "Other"},
    { value: "25", label : "Alzheimer's disease"},
    { value: "26", label : "Bell's palsy"},
    { value: "27", label : "Cerebral palsy"},
    { value: "28", label : "Epilepsy"},
    { value: "29", label : "Motor neurone disease"},
    { value: "30", label : "Multiple sclerosis"},
    { value: "31", label : "Neurofibromatosis"},
    { value: "32", label : "Parkinson's disease"},
    { value: "33", label : "Sciatica"},
    { value: "34", label : "Shingles"},
    { value: "35", label : "Other"},
    { value: "36", label : "Eczema"},
    { value: "37", label : "Psoriasis"},
    { value: "38", label : "Acne"},
    { value: "39", label : "Rosacea"},
    { value: "40", label : "Ichthyosis"},
    { value: "41", label : "Vitiligo"},
    { value: "42", label : "Hives"},
    { value: "43", label : "Seborrheic dermatitis"},
    { value: "44", label : "Other"},
    { value: "45", label : "Acromegaly"},
    { value: "46", label : "Adrenal Insufficiency %26 Addison's Disease"},
    { value: "47", label : "Cushing's Syndrome"},
    { value: "48", label : "Cystic Fibrosis"},
    { value: "49", label : "Graves' Disease"},
    { value: "50", label : "Hashimoto's Disease"},
    { value: "51", label : "Human Growth Hormone %26 Creutzfeldt-Jakob Disease"},
    { value: "52", label : "Hyperthyroidism (Overactive Thyroid)"},
    { value: "53", label : "Hypothyroidism (Underactive Thyroid)"},
    { value: "54", label : "Multiple Endocrine Neoplasia Type 1"},
    { value: "55", label : "Polycystic Ovary Syndrome (PCOS)"},
    { value: "56", label : "Pregnancy %26 Thyroid Disease"},
    { value: "57", label : "Primary Hyperparathyroidism"},
    { value: "58", label : "Prolactinoma"},
    { value: "59", label : "Thyroid Tests"},
    { value: "60", label : "Turner Syndrome"},
    { value: "61", label : "Other"},
    { value: "62", label : "Cholesteatoma"},
    { value: "63", label : "Dizziness"},
    { value: "64", label : "Dysphagia (Difficulty Swallowing)"},
    { value: "65", label : "Ear Infection (Otitis Media)"},
    { value: "66", label : "Gastric Reflux"},
    { value: "67", label : "Hearing Aids"},
    { value: "68", label : "Hearing Loss"},
    { value: "69", label : "Hoarseness"},
    { value: "70", label : "Meniere’s"},
    { value: "71", label : "Nosebleeds"},
    { value: "72", label : "Sinus Problems"},
    { value: "73", label : "Sleep Apnea"},
    { value: "74", label : "Snoring"},
    { value: "75", label : "Swimmer’s Ear (Otitis Externa)"},
    { value: "76", label : "Tinnitus (Ringing in the Ears)"},
    { value: "77", label : "Tonsils %26 Adenoid Problems"},
    { value: "78", label : "Other"},
    { value: "79", label : "Acid Reflux, Heartburn, GERD"},
    { value: "80", label : "Dyspepsia/Indigestion"},
    { value: "81", label : "Nausea and Vomiting"},
    { value: "82", label : "Peptic Ulcer Disease"},
    { value: "83", label : "Abdominal Pain Syndrome"},
    { value: "84", label : "Belching, Bloating, Flatulence"},
    { value: "85", label : "Biliary Tract Disorders, Gallbladder Disorders and Gallstone Pancreatitis"},
    { value: "86", label : "Gallstone Pancreatitis"},
    { value: "87", label : "Gallstones in Women"},
    { value: "88", label : "Constipation and Defecation Problems"},
    { value: "89", label : "Diarrhea (acute)"},
    { value: "90", label : "Diarrhea (chronic)"},
    { value: "91", label : "Irritable Bowel Syndrome"},
    { value: "92", label : "Hemorrhoids and Other Anal Disorders"},
    { value: "93", label : "Rectal Problems in Women"},
    { value: "94", label : "Other"},
    { value: "95", label : "Anemia"},
    { value: "96", label : "Hemophilia"},
    { value: "97", label : "Blood clots"},
    { value: "98", label : "Leukemia"},
    { value: "99", label : "Lymphoma"},
    { value: "100", label : "Myeloma"},
    { value: "101", label : "Other"},
    { value: "102", label : "Rheumatoid arthritis"},
    { value: "103", label : "Systemic lupus erythematosus (lupus)"},
    { value: "104", label : "Inflammatory bowel disease (IBD)"},
    { value: "105", label : "Multiple sclerosis (MS)"},
    { value: "106", label : "Type 1 diabetes mellitus"},
    { value: "107", label : "Guillain-Barre syndrome"},
    { value: "108", label : "Chronic inflammatory demyelinating polyneuropathy"},
    { value: "109", label : "Psoriasis"},
    { value: "110", label : "Graves' disease"},
    { value: "111", label : "Hashimoto's thyroiditis"},
    { value: "112", label : "Myasthenia gravis"},
    { value: "113", label : "Vasculitis"},
    { value: "114", label : "Other"},
    { value: "115", label : "Chickenpox"},
    { value: "116", label : "Common cold"},
    { value: "117", label : "Diphtheria"},
    { value: "118", label : "E. coli"},
    { value: "119", label : "Giardiasis"},
    { value: "120", label : "HIV/AIDS"},
    { value: "121", label : "Infectious mononucleosis"},
    { value: "122", label : "Influenza (flu)"},
    { value: "123", label : "Lyme disease"},
    { value: "124", label : "Malaria"},
    { value: "125", label : "Measles"},
    { value: "126", label : "Meningitis"},
    { value: "127", label : "Mumps"},
    { value: "128", label : "Poliomyelitis (polio)"},
    { value: "129", label : "Pneumonia"},
    { value: "130", label : "Rocky mountain spotted fever"},
    { value: "131", label : "Rubella (German measles)"},
    { value: "132", label : "Salmonella infections"},
    { value: "133", label : "Severe acute respiratory syndrome (SARS)"},
    { value: "134", label : "Sexually transmitted diseases"},
    { value: "135", label : "Shingles (herpes zoster)"},
    { value: "136", label : "Tetanus"},
    { value: "137", label : "Toxic shock syndrome"},
    { value: "138", label : "Tuberculosis"},
    { value: "139", label : "Viral hepatitis"},
    { value: "140", label : "West Nile virus"},
    { value: "141", label : "Whooping cough (pertussis)"},
    { value: "142", label : "Other"},
    { value: "143", label : "Acute respiratory infections"},
    { value: "144", label : "Asthma"},
    { value: "145", label : "Bronchitis"},
    { value: "146", label : "Chest pain"},
    { value: "147", label : "Diabetes"},
    { value: "148", label : "Fatigue"},
    { value: "149", label : "High blood cholesterol and triglycerides"},
    { value: "150", label : "Hypertension (high blood pressure)"},
    { value: "151", label : "Hypothyroidism"},
    { value: "152", label : "Influenza"},
    { value: "153", label : "Menopause"},
    { value: "154", label : "Migraine"},
    { value: "155", label : "Osteoarthritis"},
    { value: "156", label : "Osteoporosis"},
    { value: "157", label : "Pneumonia"},
    { value: "158", label : "Other"},
    { value: "159", label : "Familial hypercholesterolemia"},
    { value: "160", label : "Gaucher disease"},
    { value: "161", label : "Hunter syndrome"},
    { value: "162", label : "Krabbe disease"},
    { value: "163", label : "Maple syrup urine disease"},
    { value: "164", label : "Metachromatic leukodystrophy"},
    { value: "165", label : "Mitochondrial encephalopathy, lactic acidosis, stroke-like episodes (MELAS)"},
    { value: "166", label : "Niemann-Pick"},
    { value: "167", label : "Phenylketonuria (PKU)"},
    { value: "168", label : "Porphyria"},
    { value: "169", label : "Tay-Sachs disease"},
    { value: "170", label : "Wilson's disease"},
    { value: "171", label : "Other"},
    { value: "172", label : "Chronic kidney disease"},
    { value: "173", label : "Kidney stones"},
    { value: "174", label : "Glomerulonephritis"},
    { value: "175", label : "Polycystic kidney disease"},
    { value: "176", label : "Urinary tract infections"},
    { value: "177", label : "Other"},
    { value: "178", label : "Bladder Cancer"},
    { value: "179", label : "Breast Cancer"},
    { value: "180", label : "Cervical cancer"},
    { value: "181", label : "Colorectal Cancer"},
    { value: "182", label : "Kidney Cancer"},
    { value: "183", label : "Lung Cancer - Non-Small Cell"},
    { value: "184", label : "Lymphoma - Non-Hodgkin"},
    { value: "185", label : "Melanoma"},
    { value: "186", label : "Oral and Oropharyngeal Cancer"},
    { value: "187", label : "Ovarian cancer"},
    { value: "188", label : "Pancreatic Cancer"},
    { value: "189", label : "Prostate Cancer"},
    { value: "190", label : "Thyroid Cancer"},
    { value: "191", label : "Uterine Cancer"},
    { value: "192", label : "Other"},
    { value: "193", label : "Refractive Errors"},
    { value: "194", label : "Age-Related Macular Degeneration."},
    { value: "195", label : "Cataract"},
    { value: "196", label : "Diabetic Retinopathy"},
    { value: "197", label : "Glaucoma"},
    { value: "198", label : "Amblyopia"},
    { value: "199", label : "Strabismus"},
    { value: "200", label : "Other"},
    { value: "201", label : "Arthritis"},
    { value: "202", label : "Bursitis"},
    { value: "203", label : "Fibromyalgia"},
    { value: "204", label : "Foot Pain and Problems"},
    { value: "205", label : "Fractures"},
    { value: "206", label : "Low Back Pain"},
    { value: "207", label : "Hand Pain and Problems"},
    { value: "208", label : "Knee Pain and Problems"},
    { value: "209", label : "Kyphosis"},
    { value: "210", label : "Neck Pain and Problems"},
    { value: "211", label : "Osteoporosis"},
    { value: "212", label : "Paget's Disease of the Bone"},
    { value: "213", label : "Scoliosis"},
    { value: "214", label : "Shoulder Pain and Problems"},
    { value: "215", label : "Soft-Tissue Injuries"},
    { value: "216", label : "Other"},
    { value: "217", label : "Asthma"},
    { value: "218", label : "Bronchiectasis"},
    { value: "219", label : "Bronchitis"},
    { value: "220", label : "Chronic obstructive pulmonary disease (COPD)"},
    { value: "221", label : "Chronic bronchitis"},
    { value: "222", label : "Emphysema"},
    { value: "223", label : "Interstitial lung disease"},
    { value: "224", label : "Occupational lung disease"},
    { value: "225", label : "Pulmonary fibrosis"},
    { value: "226", label : "Rheumatoid lung disease"},
    { value: "227", label : "Sarcoidosis"},
    { value: "228", label : "Other"},
    { value: "229", label : "Autism Spectrum Disorder (ASD)"},
    { value: "230", label : "Schizophrenia"},
    { value: "231", label : "Bipolar Disorder"},
    { value: "232", label : "Obsessive Compulsive Disorder (OCD)"},
    { value: "233", label : "Anxiety Disorders"},
    { value: "234", label : "Phobias"},
    { value: "235", label : "Substance Use Disorder"},
    { value: "236", label : "Eating Disorders"},
    { value: "237", label : "Personality Disorders"},
    { value: "238", label : "Mood Disorders"},
    { value: "239", label : "Other"},
    { value: "240", label : "Asthma"},
    { value: "241", label : "Chronic Obstructive Pulmonary Disease (COPD)"},
    { value: "242", label : "Chronic Bronchitis"},
    { value: "243", label : "Emphysema"},
    { value: "244", label : "Lung Cancer"},
    { value: "245", label : "Cystic Fibrosis/Bronchiectasis"},
    { value: "246", label : "Pneumonia"},
    { value: "247", label : "Pleural Effusion"},
    { value: "248", label : "Other"},
    { value: "249", label : "Osteoarthritis"},
    { value: "250", label : "Rheumatoid arthritis"},
    { value: "251", label : "Lupus"},
    { value: "252", label : "Ankylosing spondylitis"},
    { value: "253", label : "Psoriatic arthritis"},
    { value: "254", label : "Sjogren’s syndrome"},
    { value: "255", label : "Gout"},
    { value: "256", label : "Scleroderma"},
    { value: "257", label : "Infectious arthritis"},
    { value: "258", label : "Juvenile idiopathic arthritis"},
    { value: "259", label : "Polymyalgia rheumatica"},
    { value: "260", label : "Other"},
    { value: "261", label : "Deep vein thrombosis"},
    { value: "262", label : "Paget-Schroetter disease"},
    { value: "263", label : "Budd-Chiari syndrome"},
    { value: "264", label : "Portal vein thrombosis"},
    { value: "265", label : "Renal vein thrombosis"},
    { value: "266", label : "Cerebral venous sinus thrombosis"},
    { value: "267", label : "Jugular vein thrombosis"},
    { value: "268", label : "Cavernous sinus thrombosis"},
    { value: "269", label : "Stroke"},
    { value: "270", label : "Myocardial infarction"},
    { value: "271", label : "Other"},
    { value: "272", label : "Bladder Cancer"},
    { value: "273", label : "Lower urinary tract symptoms."},
    { value: "274", label : "Penile and testicular cancer"},
    { value: "275", label : "Prostate cancer"},
    { value: "276", label : "Renal cancer"},
    { value: "277", label : "Urinary incontinence"},
    { value: "278", label : "Urinary tract infection"},
    { value: "279", label : "Other"},
    { value: "280", label : "Gynaecological conditions"},
    { value: "281", label : "Pregnancy conditions"},
    { value: "282", label : "Infertility disorders"},
    { value: "283", label : "Turner syndrome"},
    { value: "284", label : "Rett syndrome"},
    { value: "285", label : "Menopause"},
    { value: "286", label : "Osteoporosis"},
    { value: "287", label : "Other"}
  ];

  gates: Gate[] =  [{
    title: "Gate 1",
    opportunityId: 67,
    name: "Gate 1",
    reviewedAt: new Date("5/1/2021"),
    id: 29,
    createdAt: new Date("5/25/2021 2:45 PM"),
    actions: [],
    folders: []
  },{
    title: "Gate 2",
    opportunityId: 67,
    name: "Gate 2",
    reviewedAt: new Date("7/1/2021"),
    id: 30,
    createdAt: new Date("5/25/2021 3:04 PM"),
    actions: [],
    folders: []
  },{
    title: "Gate 1",
    opportunityId: 68,
    name: "Gate 1",
    reviewedAt: new Date("4/1/2022"),
    id: 31,
    createdAt: new Date("5/25/2021 3:55 PM"),
    actions: [],
    folders: []
  },{
    title: "Phase 1",
    opportunityId: 69,
    name: "Phase 1",
    reviewedAt: new Date("1/1/2022"),
    id: 32,
    createdAt: new Date("5/25/2021 3:58 PM"),
    actions: [],
    folders: []
  },{
    title: "Phase 2",
    opportunityId: 69,
    name: "Phase 2",
    reviewedAt: new Date("12/1/2022"),
    id: 33,
    createdAt: new Date("5/25/2021 4:01 PM"),
    actions: [],
    folders: []
  },{
    title: "Phase 3",
    opportunityId: 69,
    name: "Phase 3",
    reviewedAt: new Date("6/1/2023"),
    id: 34,
    createdAt: new Date("5/25/2021 4:02 PM"),
    actions: [],
    folders: []
  }];

  actions: Action[] = [{
    gateId: 29,
    id: 1,
    opportunityId: 67,
    title:  "Commercial terms negotiations",
    actionName:  "Commercial terms negotiations",
    dueDate: new Date("4/29/2021"),
    completed: true,
    timestamp: new Date("6/7/2021 11:43 AM"),
    targetUserId: 1,
    targetUser: {
      id: 1,
      name: "David Baker"
    }
  },{
    gateId: 29,
    id: 2,
    opportunityId: 67,
    title:  "Innovation board",
    actionName:  "Innovation board",
    dueDate: new Date("3/5/2021"),
    completed: true,
    timestamp: new Date("6/7/2021 11:43 AM"),
    targetUserId: 1,
    targetUser: {
      id: 1,
      name: "David Baker"
    }
  },{
    gateId: 29,
    id: 3,
    opportunityId: 67,
    title:  "SMT Approval",
    actionName:  "SMT Approval",
    dueDate: new Date("4/5/2021"),
    completed: true,
    timestamp: new Date("6/7/2021 11:43 AM"),
    targetUserId: 1,
    targetUser: {
      id: 1,
      name: "David Baker"
    }
  },{
    gateId: 29,
    id: 4,
    opportunityId: 67,
    title:  "DD/Contract approving process",
    actionName:  "DD/Contract approving process",
    dueDate: new Date("5/5/2021"),
    completed: true,
    timestamp: new Date("6/7/2021 11:43 AM"),
    targetUserId: 1,
    targetUser: {
      id: 1,
      name: "David Baker"
    }
  },{
    gateId: 29,
    id: 5,
    opportunityId: 67,
    title:  "Commercial terms negotiations",
    actionName:  "Commercial terms negotiations",
    dueDate: new Date("6/29/2021"),
    completed: true,
    timestamp: new Date("6/7/2021 11:43 AM"),
    targetUserId: 1,
    targetUser: {
      id: 1,
      name: "David Baker"
    }
  },{
    gateId: 29,
    id: 6,
    opportunityId: 67,
    title:  "Innovation board",
    actionName:  "Innovation board",
    dueDate: new Date("7/5/2021"),
    completed: false,
    timestamp: new Date("6/7/2021 11:43 AM"),
    targetUserId: 1,
    targetUser: {
      id: 1,
      name: "David Baker"
    }
  },{
    gateId: 29,
    id: 7,
    opportunityId: 67,
    title:  "SMT Approval",
    actionName:  "SMT Approval",
    dueDate: new Date("8/5/2021"),
    completed: false,
    timestamp: new Date("6/7/2021 11:43 AM"),
    targetUserId: 1,
    targetUser: {
      id: 1,
      name: "David Baker"
    }
  },{
    gateId: 29,
    id: 8,
    opportunityId: 67,
    title:  "DD/Contract approving process",
    actionName:  "DD/Contract approving process",
    dueDate: new Date("9/5/2021"),
    completed: false,
    timestamp: new Date("6/7/2021 11:43 AM"),
    targetUserId: 1,
    targetUser: {
      id: 1,
      name: "David Baker"
    }
  }];

  /*
"Registration changes (MA owner)","Gate 2","Acquisition of Nucala for COPD","Registration changes (MA owner)","4/5/2021","Sí","7/6/2021 9:04 AM","Marc Torruella Altadill"
"QA Audit","Gate 2","Acquisition of Nucala for COPD","QA Audit","5/5/2021","Sí","7/6/2021 9:04 AM","Marc Torruella Altadill"
"Contract signing","Gate 2","Acquisition of Nucala for COPD","Contract signing","6/4/2021","Sí","7/6/2021 9:04 AM","Marc Torruella Altadill"
"Registration changes (MA owner)","Gate 2","Acquisition of Nucala for COPD","Registration changes (MA owner)","7/4/2021","Sí","7/6/2021 9:04 AM","Marc Torruella Altadill"
"Commercial terms negotiations","Gate 1","Acquisition of Tezepelumab (Asthma)","Commercial terms negotiations","2/4/2021","No","6/8/2021 4:55 PM","Marc Torruella Altadill"
"Innovation board","Gate 1","Acquisition of Tezepelumab (Asthma)","Innovation board","3/6/2021","Sí","6/8/2021 4:55 PM","Marc Torruella Altadill"
"SMT Approval","Gate 1","Acquisition of Tezepelumab (Asthma)","SMT Approval","4/5/2021","No",,
"DD/Contract approving process","Gate 1","Acquisition of Tezepelumab (Asthma)","DD/Contract approving process","5/5/2021","No",,
"Commercial terms negotiations","Gate 1","Acquisition of Tezepelumab (Asthma)","Commercial terms negotiations","6/4/2021","No",,
"Innovation board","Gate 1","Acquisition of Tezepelumab (Asthma)","Innovation board","7/4/2021","No",,
"SMT Approval","Gate 1","Acquisition of Tezepelumab (Asthma)","SMT Approval","8/3/2021","No",,
"DD/Contract approving process","Gate 1","Acquisition of Tezepelumab (Asthma)","DD/Contract approving process","9/2/2021","No",,
"Initiation and Prototyping (incl API sourcing and decision making)","Phase 1","Development of Concizumab","Initiation and Prototyping (incl API sourcing and decision making)","2/4/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Formulation optimisation","Phase 1","Development of Concizumab","Formulation optimisation","3/6/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Pre-Clinical study (with Report)","Phase 1","Development of Concizumab","Pre-Clinical study (with Report)","4/5/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Pilot BE (incl CTA and supplies)","Phase 1","Development of Concizumab","Pilot BE (incl CTA and supplies)","5/5/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Final Business case","Phase 1","Development of Concizumab","Final Business case","6/4/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Tech Transfer","Phase 2","Development of Concizumab","Tech Transfer","2/4/2021","Sí","5/25/2021 4:02 PM","David Baker"
"Stability (Regulatory batches)","Phase 2","Development of Concizumab","Stability (Regulatory batches)","3/6/2021","Sí","5/25/2021 4:02 PM","David Baker"
"Pivotal BE study (incl CTA, supplies and CSR)","Phase 2","Development of Concizumab","Pivotal BE study (incl CTA, supplies and CSR)","4/5/2021","Sí","5/25/2021 4:02 PM","David Baker"
"Phase III Clinical study (incl CTA, supplies and CSR)","Phase 3","Development of Concizumab","Phase III Clinical study (incl CTA, supplies and CSR)","2/4/2021","No",,
"Market Authorisation Submission-Approval","Phase 3","Development of Concizumab","Market Authorisation Submission-Approval","3/6/2021","No",,
"Patent expiry","Phase 3","Development of Concizumab","Patent expiry","4/5/2021","No",,
"Launch activities (including pricing/reimbursement)","Phase 3","Development of Concizumab","Launch activities (including pricing/reimbursement)","5/5/2021","No",,

  */
  constructor(private teams: TeamsService, private http: HttpClient, private error: ErrorService, private licensing: LicensingService) { }

  query(url: string) {
    return this.http.get(this.licensing.siteUrl + url, { headers: this.buildDefaultHeaders() });
  }

  buildDefaultHeaders(): any {
    let headersObject = new HttpHeaders({
      'Accept':'application/json;odata=verbose',
      'Authorization': 'Bearer ' + this.teams.token
    });
    return headersObject;
  }

  async getOpportunities(): Promise<Opportunity[]> {
    return this.opportunities;
  }

  async getIndications(): Promise<any[]> {
    return this.indications;
  }

  async getGates(opportunityId: number): Promise<Gate[]> {
    return this.gates.filter(el => el.opportunityId == opportunityId);
  }

  async getActions(gateId: number): Promise<Action[]> {
    return this.actions.filter(el => el.gateId == gateId);
  }

  async getLists() {
   /* try {
      let lists = await this.query('lists').toPromise();
      return lists;
    } catch (e) {
      if(e.status == 401) {
        this.teams.loginAgain();
      }
      return [];
    }*/
  }
  async getOpportunityTypes() {
    return [
      { value: 'acquisition', label: 'Acquisition' },
      { value: 'licensing', label: 'Licensing' },
      { value: 'productDevelopment', label: 'Product Development' }
    ];
  }

  async getOpportunityFields() {
    return [
      { value: 'title', label: 'Opportunity Name' },
      { value: 'projectStart', label: 'Project Start Date' },
      { value: 'projectEnd', label: 'Project End Date' },
      { value: 'opportunityType', label: 'Project Type' },
    ];
  }

  async getOpportunity(id: number) {
    return this.opportunities.find(el => el.Id == id);
  }

  async getFiles(id: number) {
    return this.files.filter(f => f.parentId == id);
  }
}
