// 스프레드시트 ID를 여기에 붙여넣으세요.
const MAIN_SPREADSHEET_ID = '1uMGEUnktPuj8nwFdni-PKGnoZrwhaChBRFpKPTW_a3M';
const SHEET_NAME = 'Sheet1';
const DRIVE_FOLDER_ID = '1jc201O_cmnZ0yz3vHc00YrWtxYc9m5fj';

// Office별 스프레드시트 ID 매핑
const OFFICE_SPREADSHEET_IDS = {
  'Ming': '1zEiT8r7Pt2XVs2aK_JVSX5cs1Cczxe1OLGIo8PtjHQo',
  'Bernard': '1hTZj-uSocMpe1HQXGJ9ERT5gvoq6SsVN-SXdaJgZijo',
  'Delano': '1zbcuCd1EhsGJ8__UJJTGl3B6j4iaw3aFpU8RTZOAjqQ',
  'Tulare': '1f3rUJjUA69Oz2G0YsbcFBBsmx-lqfDbYyWMtRPr-7mk',
  'Visalia': '1RfGeHZH1n74ge2lfrFpARd8QIOIFihEAWSfOyXiScjc',
  'Fresno': '1n_OoR2OgXeOR5z4UdiLCdJWbk6jrt_8Esvvs1qdYhXg',
  'California': '1wLwFe2n-NKm9xq60CSPXZJUzn1i5v_Tl2mExq7BeVH8',
  'Ortho': '1Ozjt3MZgQ_tggqi4OP8mMN9xQynwrOniKlIL1mqiGjQ',
  'EXEC': '1FMu0cYowZRxvNderIS8FrtsuxLp0cuJp7G4EP5yeqRI',
  'AR': '1gY1FyVHS-YBqdp8y3S3CBdxd22BtvBj-RT1TeWZlGJc',
  'Call Center': '1vcWLP5glRmG6yyXuLklOMiN1myD34ZRG_JyOYji-kik',
  'IT': '13MWqBI2XVR71gAzHPiKpjEFfwWD-8jSNeyeVzXKq6XA',
  'Marcom': '1PVv96EDYO-d53w1mptfTLSxvBps4obQThOpLpKIzIjE',
  'Office Billing': '1giCxaGuLTFc3z_j30AQkLIpmnrQH9VqraZdgayNXiDE',
};

// Office별 이메일 주소 매핑
const OFFICE_EMAILS = {
  'Ming': 'kroyal@smlnd.com',
  'Bernard': 'bk1.bern@smlnd.com',
  'Delano': 'bk.delano@smlnd.com',
  'Tulare': 'drojas@smlnd.com',
  'Visalia': 'drojas@smlnd.com',
  'Fresno': 'cromero@smlnd.com',
  'California': 'kroyal@smlnd.com',
  'Ortho': 'sj@smlnd.com',
  'Office Billing': 'kmorillo@smlnd.com'
};

// Corporate 오피스 상태별 캘린더 ID 매핑
const CORPORATE_STATUS_CALENDAR_IDS = {
  'approved': 'c_b2bab8d9b31c3f447a89dfd783f1277be7ebc5e3671732ac2c31e2a26bd2921f@group.calendar.google.com',
  'approved – excuse note required': 'c_f22dab34eac33ca369cb9fd3f9873671a64afe1fc1ed25ca6638a8a2699c99d6@group.calendar.google.com',
  'requested': 'c_9c81e705a567919f337b9be64598b842d239ecd021edb30c4e1ad41a3ea1eeba@group.calendar.google.com',
  'denied': 'c_f81b7d856fb1257b19dd3a0e7f56922954f40a23ada95b92b5c9b9335fe7066e@group.calendar.google.com',
  'approved - cancellation': 'c_1dfc40d1776f968938839789a5929b89f220dfe70210b6bf2e7b1e4c0e514560@group.calendar.google.com',
  'pending': 'c_334241937c2cdb995099c7c363f0ade6915dbfa9e5bc154416529ac3449e8dd4@group.calendar.google.com',
  'called in': 'c_3963a75d17dc58a7d9058d3352313e02d7c8492aca5adb6c62eaff2fa217f8e9@group.calendar.google.com'
};

// Corporate 부서별 상태별 캘린더 ID 매핑 (예시: 실제 ID로 교체 필요)
const CORPORATE_DEPARTMENT_STATUS_CALENDAR_IDS = {
  'AR': {
  'approved': 'c_6c63cf38a32e544a303599a2a11bb9eeb55a4fbbb1110509acb17db3e70dea9b@group.calendar.google.com',
  'approved – excuse note required': 'c_41535513da644255726e1f6863cb5eb94a84aeaf9e89baecdd5bf9102eccfa90@group.calendar.google.com',
  'approved - cancellation': 'c_ea74f0f67772ccb88383227f08ccd1f12fcbc46f112f231a0efe18ed6b4ad74a@group.calendar.google.com',
  'denied': 'c_c29f010dc6e4b4d22f151b7d053f6c68404b98d8de44f017d9d45aefa32a24bb@group.calendar.google.com',
  'pending': 'c_8115a38872b8227bb4fb94453fe6bd958757d90540429f4e46045789745ba548@group.calendar.google.com',
  'requested': 'c_76a49bcbce32dec0b8e8249b850f64166af7f5e71cf22b8c120db2677e5c174a@group.calendar.google.com',
  'cancelled': 'c_ea74f0f67772ccb88383227f08ccd1f12fcbc46f112f231a0efe18ed6b4ad74a@group.calendar.google.com'
  },
  'Call Center': {
    'approved': 'c_5deaa5958a24ca0919502f54b366ee877abbf9458d60bd182c30b183ee9a7081@group.calendar.google.com',
    'approved – excuse note required': 'c_7b6deafb56146683fe62d650b6ee9d933a429890ef2392e5e3739cc395fc5208@group.calendar.google.com',
    'approved - cancellation': 'c_422cb83cc32a137d26a2bca0f34c7f48d21e66cf99830c712bac0b663cc27ac8@group.calendar.google.com',
    'denied': 'c_faadaed26a49e7acc715dbc6586d3a9c7d33e60893f2038e8ac95fa878188c26@group.calendar.google.com',
    'pending': 'c_58d2d63a8aa6c8b81d92b57688c784f47bdfa7929ec764dce1d167c249a609a7@group.calendar.google.com',
    'requested': 'c_98d835716933f6e292a1ca9ff0ee6de10976621749d94ecdaa81448dacf5d94e@group.calendar.google.com',
    'cancelled': 'c_422cb83cc32a137d26a2bca0f34c7f48d21e66cf99830c712bac0b663cc27ac8@group.calendar.google.com'
  },
  'EXEC': {
    'approved': 'c_b2bab8d9b31c3f447a89dfd783f1277be7ebc5e3671732ac2c31e2a26bd2921f@group.calendar.google.com',
    'approved – excuse note required': 'c_f22dab34eac33ca369cb9fd3f9873671a64afe1fc1ed25ca6638a8a2699c99d6@group.calendar.google.com',
    'approved - cancellation': 'c_1dfc40d1776f968938839789a5929b89f220dfe70210b6bf2e7b1e4c0e514560@group.calendar.google.com',
    'denied': 'c_f81b7d856fb1257b19dd3a0e7f56922954f40a23ada95b92b5c9b9335fe7066e@group.calendar.google.com',
    'pending': 'c_334241937c2cdb995099c7c363f0ade6915dbfa9e5bc154416529ac3449e8dd4@group.calendar.google.com',
    'requested': 'c_9c81e705a567919f337b9be64598b842d239ecd021edb30c4e1ad41a3ea1eeba@group.calendar.google.com',
    'cancelled': 'c_1dfc40d1776f968938839789a5929b89f220dfe70210b6bf2e7b1e4c0e514560@group.calendar.google.com'
  },
  'IT': {
    'approved': 'c_2bc51def963ccad44a7097ff1645614e67c190e2ca8f97a8f71ba1f12020d779@group.calendar.google.com',
    'approved – excuse note required': 'c_62adcf074fef4ce141a79db9abbd6f2d460d9fe3abaa32ce311271cc2b00111c@group.calendar.google.com',
    'approved - cancellation': 'c_862b4b4b041e8c81e8789183afcb58918c269e1521c448fc0d20ca943ccaa16a@group.calendar.google.com',
    'denied': 'c_660908384f64250ed5036b9961b5a4c43dd3823724f0253d97582385ffa6a6f6@group.calendar.google.com',
    'pending': 'c_5c35cad0d35da1b00a10e74fbc2ce8799d80de878b3b73ce6e993c2bec1e221b@group.calendar.google.com',
    'requested': 'c_8b54bbc71d2cdca10dd3667d1e23fb5b1528c4852244765bc463d86ed12d5030@group.calendar.google.com',
    'cancelled': 'c_862b4b4b041e8c81e8789183afcb58918c269e1521c448fc0d20ca943ccaa16a@group.calendar.google.com'
  },
  'Marcom': {
    'approved': 'c_14c3c3f0b475dcf48ae4e2258542f5b5e10b041273c0b943f01212ec6ae14ee1@group.calendar.google.com',
    'approved – excuse note required': 'c_f3085f78d82f32b1403292aa73dc97e5c9f195346066d24da97fe8deb8919c6d@group.calendar.google.com',
    'approved - cancellation': 'c_ba771d1ae2f157f42c2d280bd0fed9363a6d123953ce8e02e1870b4d3b67647f@group.calendar.google.com',
    'denied': 'c_2e2e890c6a2d7d71c5e09b980f17b38e502551e1eee3508178503574fdfbf239@group.calendar.google.com',
    'pending': 'c_e3719a02c82eb361a7a485af9523a46932798e60a453d251eaeb4a533fa53b73@group.calendar.google.com',
    'requested': 'c_923e6e904b5dce85981f0c0994023d6e3b54413e28cf393ef7bda38501d6f80f@group.calendar.google.com',
    'cancelled': 'c_ba771d1ae2f157f42c2d280bd0fed9363a6d123953ce8e02e1870b4d3b67647f@group.calendar.google.com'
  }
};

// Corporate 부서별 이메일 주소 매핑
const CORPORATE_DEPARTMENT_EMAILS = {
  'EXEC': 'mruiz@smlnd.com',
  'Marcom': 'jhkim@smlnd.com',
  'AR': 'atovar@smlnd.com',
  'Call Center': 'gmedina@smlnd.com',
  'IT': 'jchoe@smlnd.com'
};

// Ming 오피스 상태별 캘린더 ID 매핑 추가
const MING_STATUS_CALENDAR_IDS = {
  'approved': 'c_7d64a5a158be5a291adda62268ca60f9666c4631d887a0030cd4ab748b1bf91f@group.calendar.google.com',
  'approved – excuse note required': 'c_c79d26b24fdeb162a034cc9a7cddd3623083002e5a43ebf85901d954f96fa157@group.calendar.google.com',
  'requested': 'c_d56dbfb55494eee8a860fc3eabae602ea5687ec3b13d6c1685b28dcab0a54297@group.calendar.google.com',
  'denied': 'c_bfb6b9f43c6bbdf64000f07fcad02f1441c9c6b1184aec9f7d645e4d7abfdd1d@group.calendar.google.com',
  'called in': 'c_d808a68d9750f96698729bd0baa3ed8155c426a4be9ccd600e5c0946a16c6f60@group.calendar.google.com',
  'cancelled': 'c_52280d0e1aab27f73744a9dc13c5260eaf75e97e4454490be62c16335c02af0a@group.calendar.google.com',
  'pending': 'c_e3c32017dc724f4457d553e4ab6d241ee6b859d0e8c8feeaa646d8bbd0f97349@group.calendar.google.com'
};

// Bernard 오피스 상태별 캘린더 ID 매핑 추가
const BERNARD_STATUS_CALENDAR_IDS = {
  'approved': 'c_8d8bb1a92ec0e221e020216f339b2d1202fe4bf8f95596b8e5cf57b1d88f01af@group.calendar.google.com',
  'approved – excuse note required': 'c_c91efb272498b7275cd010a7a515ce8f692c03089cc9a9fe3db95329d4f1b126@group.calendar.google.com',
  'denied': 'c_3aff0504c75d8ffd99c831110484e0f6d0a67140f9c68e512d964544015c2be9@group.calendar.google.com',
  'called in': 'c_d1f9038f5ff5311f8ca2a97850e02342604a7c07e009db6e604f070a7794c3e1@group.calendar.google.com',
  'cancelled': 'c_0281c4bb0528fb27da462d04310a336123459240259d01109978373e3bdbdb67@group.calendar.google.com',
  'pending': 'c_05af3a0b60007e9751697bec973496e3297566d05d4f974d2f82e97caec4ed38@group.calendar.google.com',
  'requested': 'c_7a7d5a2ce6ce92a501ba82ecd42ada79ac09f6831529727fa1287741203dbbe8@group.calendar.google.com'
};

// Delano 오피스 상태별 캘린더 ID 매핑 추가
const DELANO_STATUS_CALENDAR_IDS = {
  'approved': 'c_37c3847635f73c89b5b41da527f494779de1a260bcdb4b30b5477787e3f520b1@group.calendar.google.com',
  'approved – excuse note required': 'c_2c9eadfa94e301444930ff333264a01f7d978e537bbae5e5896afe46d42cde3d@group.calendar.google.com',
  'denied': 'c_2a9818ebfef62eac79c6535662d98a7499a294f048e2d7bc331616ee90fe11a9@group.calendar.google.com',
  'called in': 'c_c4a65616e751f542a9e90974418be462794badc9aa250e3f37d8a7a0889704f4@group.calendar.google.com',
  'cancelled': 'c_be58ceb1a24684bcfd22ab196b1754a9197292969a8e481925b65e38e72c117b@group.calendar.google.com',
  'pending': 'c_fbb200697888c0ec00d491bc5a43b5e4189ef21a90f539dd7b0cc03d6d918b73@group.calendar.google.com',
  'requested': 'c_5d3929402bb6ff299f4e31e971f02eb663985a83af9cebd39f771cd6bdd92d77@group.calendar.google.com'
};

// Tulare 오피스 상태별 캘린더 ID 매핑 추가
const TULARE_STATUS_CALENDAR_IDS = {
  'approved': 'c_9be501444cf8987f57a7f94e3b30094f34f3764359371746423355c5bbb142c3@group.calendar.google.com',
  'approved – excuse note required': 'c_4ebd71987692a48c7560847119d8b5a6e17348a87e9ecf916e06e501e98950d4@group.calendar.google.com',
  'denied': 'c_9bb2a3e6ce8882994fb01c85910b221b764a3aaca01ee6119d122c96f4e404b0@group.calendar.google.com',
  'called in': 'c_a48409999f3830e4db43461430e22c3cc4fc02e65570f0553631a42a73d69f3b@group.calendar.google.com',
  'cancelled': 'c_bd954af0b87979c05198c7b9328cf4c88af0e369b1d2172d5213f90d05f28585@group.calendar.google.com',
  'pending': 'c_3e9989d00206e9073bacc5fcc8a6808888c05a6a8259a8faddf78154e9493546@group.calendar.google.com',
  'requested': 'c_b28b6bd8669b4759a6134048a90c34204864c6bc65ca4af5dbe7bc4b8b5d9d23@group.calendar.google.com'
};

// Visalia 오피스 상태별 캘린더 ID 매핑 추가
const VISALIA_STATUS_CALENDAR_IDS = {
  'approved': 'c_ba51e791d8253c7f965e7cac59ad458b2b6572c7c2f8f8580f23f9f295cf7864@group.calendar.google.com',
  'approved – excuse note required': 'c_2dcbee6c02a4f7500ac89282a7daf8a402a32275038d5b6f4d89333bbfccd66e@group.calendar.google.com',
  'denied': 'c_92e236cd3016a62e843750ffc60f2bfe96f17f181709a8330e6bcdda34e1b8a6@group.calendar.google.com',
  'called in': 'c_9a68e6d04e47cd4b7e179e87e0c33e3aba4b65344f2d9d1aaff386ac437da0da@group.calendar.google.com',
  'cancelled': 'c_7a5220171e2e7300fc36b8986a411d701a0fd55741e3d6973888f7a63c0c9b76@group.calendar.google.com',
  'pending': 'c_ede5731e4bd251e521bc0900ae512fb1ceb5d1f26674adbc519fe11d41f61509@group.calendar.google.com',
  'requested': 'c_8842265b235bda3a2a7bdb4c278022d978923438f254da7f7ef1f59491cd9329@group.calendar.google.com'
};

// Fresno 오피스 상태별 캘린더 ID 매핑 추가
const FRESNO_STATUS_CALENDAR_IDS = {
  'approved': 'c_9eda21903469650aebe9732e7564683ace61baf9840f8a53588762f611a89b02@group.calendar.google.com',
  'approved – excuse note required': 'c_53d8f2994c587b66b9ea92d684d032adb29a12bc2b2cf58412e5c4e850774984@group.calendar.google.com',
  'denied': 'c_6156f58be73f3a6ac420ad095377dd4c2a600cade40fd391ec0a14a9a7d2e673@group.calendar.google.com',
  'called in': 'c_abfb616558fd05bd7f67a0a7f665575ed578dfd98c48600f3b4a02786a9c9057@group.calendar.google.com',
  'cancelled': 'c_3441258b25069bafd75186a6046a3ffabb4c72d3cba6601da61ac23be24d9dd1@group.calendar.google.com',
  'pending': 'c_b75d04f8acf23b4487622c8891aff282afe3de18420b37316bf92d46bd09e29f@group.calendar.google.com',
  'requested': 'c_f55f9a7dde5578f1e02b5d81e7c08a5462260025ec97154aeab962aa795ef7d5@group.calendar.google.com'
};

// California 오피스 상태별 캘린더 ID 매핑 추가
const CALIFORNIA_STATUS_CALENDAR_IDS = {
  'approved': 'c_1d389626fdac6e4b6ffd42da354d5a54bf95629f488298d935cc7eaf60614086@group.calendar.google.com',
  'approved – excuse note required': 'c_79b6e67d107dbb5909329c7afe8c34df4bd699a1c584ba36bf2403ea402276cf@group.calendar.google.com',
  'denied': 'c_3fd9757ef06d65354ae9dfeaa036b6051b59e42909b11d571cdcc540777babc7@group.calendar.google.com',
  'called in': 'c_11f8e05d14fca3f373a64281d51d4d6846ec24ef41c13c9e169e8cf0a855b58a@group.calendar.google.com',
  'cancelled': 'c_e29560f0d96285b4179dc1e10d383ef39e40937a06f7e4044b56d24b05390be0@group.calendar.google.com',
  'pending': 'c_17f3e10400d657a6e47ac82bc4a84a70d33d14ecf1b4238a9dce926a48cc2dfb@group.calendar.google.com',
  'requested': 'c_900b97649c9179ea0a72ba1dc6e5258318938f70f38fea55b7aa72c2939feebb@group.calendar.google.com'
};

// Ortho 오피스 상태별 캘린더 ID 매핑 추가
const ORTHO_STATUS_CALENDAR_IDS = {
  'approved': 'c_73f2ad5e421b4ae97c044b487087d489c84b381b4306a539f133de1fc852e5e0@group.calendar.google.com',
  'approved – excuse note required': 'c_d913c053fd1f750c73c3cfe7dcf19eeac49087fc64e7fe325078d6600d6cdbf4@group.calendar.google.com',
  'denied': 'c_a02618104f8c272fd8f80be68f463a8a3fe55d93e0a69d5cabe9b4c8a17a79f5@group.calendar.google.com',
  'called in': 'c_46516d06abd6180efa493222416046d63f1901daf717b8611cf72bbdbf0b5deb@group.calendar.google.com',
  'cancelled': 'c_eef28ca1803518e4f74207c80f8c480d2d8564c489292e22c24d6d435a0a1173@group.calendar.google.com',
  'pending': 'c_b6f0ee3e650ca6af5b8119910add942fdeba1107a7d49eabedcf13b64fb676ef@group.calendar.google.com',
  'requested': 'c_5973356a25b30b4cc74c8856bcbaf8535b296b253f8b7b5fc01f6e4e72ab4c14@group.calendar.google.com'
};

// Office Billing 오피스 상태별 캘린더 ID 매핑 추가
const OFFICE_BILLING_STATUS_CALENDAR_IDS = {
  'approved': 'c_2655fa5be1c42d1825ea3c066ab88f0ce530c40a556c221dad10896546a15d07@group.calendar.google.com',
  'approved – excuse note required': 'c_c6bde94f8d0f68b0d5b5157a236677b63a45e15e3322dfc22968337a380ec3e9@group.calendar.google.com',
  'denied': 'c_6a0847215674a31cb430b461f30581b469a0de203e1319e0b197db89553b8109@group.calendar.google.com',
  'called in': 'c_73d81a75fa5f65b25572bf9c232891f03289e19f8306ccd49f05118d558e8b29@group.calendar.google.com',
  'cancelled': 'c_ad4fe33dc9d05b87668484fd1ee7ebbc572aee636b0f5fc1bc4e7a815f543b86@group.calendar.google.com',
  'pending': 'c_98a491658632338fada518b495750c28f1f441be31ad3df81af07fd0f0d4dd6c@group.calendar.google.com',
  'requested': 'c_98a491658632338fada518b495750c28f1f441be31ad3df81af07fd0f0d4dd6c@group.calendar.google.com'
};

// Office Billing AR 부서 상태별 캘린더 ID 매핑 추가
const OFFICE_BILLING_AR_STATUS_CALENDAR_IDS = {
  'approved': 'c_6c63cf38a32e544a303599a2a11bb9eeb55a4fbbb1110509acb17db3e70dea9b@group.calendar.google.com',
  'approved – excuse note required': 'c_41535513da644255726e1f6863cb5eb94a84aeaf9e89baecdd5bf9102eccfa90@group.calendar.google.com',
  'denied': 'c_ea74f0f67772ccb88383227f08ccd1f12fcbc46f112f231a0efe18ed6b4ad74a@group.calendar.google.com',
  'called in': 'c_c29f010dc6e4b4d22f151b7d053f6c68404b98d8de44f017d9d45aefa32a24bb@group.calendar.google.com',
  'cancelled': 'c_8115a38872b8227bb4fb94453fe6bd958757d90540429f4e46045789745ba548@group.calendar.google.com',
  'pending': 'c_76a49bcbce32dec0b8e8249b850f64166af7f5e71cf22b8c120db2677e5c174a@group.calendar.google.com',
  'requested': 'c_76a49bcbce32dec0b8e8249b850f64166af7f5e71cf22b8c120db2677e5c174a@group.calendar.google.com'
};

const ALLOWED_STATUSES = ['approved', 'approved – excuse note required', 'denied', 'approved - cancellation', 'pending', 'requested', 'called in'];

// 웹앱 접속 시 보여줄 폼
function doGet(e) {
  Logger.log('doGet called');
  return HtmlService
    .createHtmlOutputFromFile('form')
    .setTitle('Absence Request/Report');
}

// HTML 파일 include 용 헬퍼
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
                    .getContent();
}

// PDF를 구글 드라이브에 저장
function savePDFToDrive(pdfBlob, fileName) {
  try {
    Logger.log('savePDFToDrive called');
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    Logger.log('Folder found: ' + folder.getName());
    const file = folder.createFile(pdfBlob);
    file.setName(fileName);
    Logger.log('File saved: ' + file.getUrl());
    return file.getUrl();
  } catch (e) {
    Logger.log('Error in savePDFToDrive: ' + e);
    throw e;
  }
}

// YYYY-MM-DD 포맷을 정확하게 Date 객체로 변환 (타임존 문제 방지)
function parseDateWithTimezone(dateString) {
  // yyyy-mm-dd 또는 mm/dd/yyyy 등 다양한 포맷 대응
  if (!dateString) return null;
  if (typeof dateString === 'object' && dateString instanceof Date) return dateString;
  // yyyy-mm-dd
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
  var parts = dateString.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }
  // mm/dd/yyyy
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateString)) {
    var parts = dateString.split('/');
    return new Date(Number(parts[2]), Number(parts[0]) - 1, Number(parts[1]));
  }
  // 기타: Date.parse 시도
  var d = new Date(dateString);
  if (!isNaN(d.getTime())) return d;
  return null;
}

// Corporate 오피스에서 부서별로 캘린더 ID를 선택하는 함수
function getCorporateCalendarId(department, status) {
  if (CORPORATE_DEPARTMENT_STATUS_CALENDAR_IDS[department] && CORPORATE_DEPARTMENT_STATUS_CALENDAR_IDS[department][status]) {
    return CORPORATE_DEPARTMENT_STATUS_CALENDAR_IDS[department][status];
  }
  // fallback: 기존 corporate calendar id
  return CORPORATE_STATUS_CALENDAR_IDS[status] || CORPORATE_STATUS_CALENDAR_IDS['requested'];
}

// 이메일 주소 결정 함수 (폼 제출 시 사용)
function getEmailAddresses(office, department) {
  let ccEmails = ['attendance@smlnd.com']; // 기본 attendance 이메일
  
  if (office === 'Corporate' && department) {
    // Corporate 오피스의 경우 부서별 이메일 추가
    const departmentEmail = CORPORATE_DEPARTMENT_EMAILS[department];
    if (departmentEmail) {
      ccEmails.push(departmentEmail);
    }
  } else if (office && OFFICE_EMAILS[office]) {
    // 일반 오피스의 경우 오피스별 이메일 추가
    ccEmails.push(OFFICE_EMAILS[office]);
  }
  
  return ccEmails.join(',');
}

// 승인 상태 변경 시 이메일 주소 결정 함수 (attendance@smlnd.com 제외)
function getEmailAddressesForStatusChange(office, department) {
  let ccEmails = [];
  
  if (office === 'Corporate' && department) {
    // Corporate 오피스의 경우 부서별 이메일 추가
    const departmentEmail = CORPORATE_DEPARTMENT_EMAILS[department];
    if (departmentEmail) {
      ccEmails.push(departmentEmail);
    }
  } else if (office && OFFICE_EMAILS[office]) {
    // 일반 오피스의 경우 오피스별 이메일 추가
    ccEmails.push(OFFICE_EMAILS[office]);
  }
  
  return ccEmails.join(',');
}

function submitRequest(formData) {
  try {
    Logger.log('submitRequest called');
    Logger.log('formData: ' + JSON.stringify(formData));
    const office = formData.office;
    // Type of Request에 따라 시트명 결정
    let sheetName = 'Sheet1';
    if (formData.requestType === 'Incident Notice') sheetName = 'Incident Notice';
    else if (formData.requestType === 'Time-Off Request') sheetName = 'Time-Off Request';
    else if (formData.requestType === 'Taken') sheetName = 'Taken';
    Logger.log('sheetName: ' + sheetName);

    const mainSs = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    Logger.log('Spreadsheet opened');
    let mainSheet = mainSs.getSheetByName(sheetName);
    Logger.log('mainSheet: ' + (mainSheet ? 'found' : 'not found'));

    // 시트가 없으면 새로 생성 (헤더도 복사)
    if (!mainSheet) {
      mainSheet = mainSs.insertSheet(sheetName);
      const headers = [
        "Type of Request", "Name", "Today's Date", "Email", "Office", "Department", "Position",
        "Type of Incident", "Cancelling Incident", "ETA", "ETD", "Start Date", "End Date",
        "Reason for Incident", "Opening", "Closing", "Excuse Note", "Excuse Note File",
        "Signature", "Schedule Conflict", "Supervisor Signature", "Send to Corp", "Status"
      ];
      mainSheet.appendRow(headers);
      Logger.log('Sheet created and headers appended');
    }

    const rowData = [
      formData.requestType,           // 0 Type of Request
      formData.name,                  // 1 Name
      formData.todayDate,             // 2 Today's Date
      formData.email,                 // 3 Email
      formData.office,                // 4 Office
      formData.department,            // 5 Department
      formData.position,              // 6 Position
      formData.incidentType,          // 7 Type of Incident
      formData.cancelTargetType,      // 8 Cancelling Incident
      formData.eta || '',             // 9 ETA
      formData.etd || '',             // 10 ETD
      formData.startDate,             // 11 Start Date
      formData.endDate,               // 12 End Date
      formData.reason,                // 13 Reason for Incident
      formData.opening || '',         // 14 Opening
      formData.closing || '',         // 15 Closing
      formData.excuseType,            // 16 Excuse Note (Submitted, Will provide, Will not provide)
      formData.excuseFile || '',      // 17 Excuse Note File (파일명 저장)
      formData.signature,             // 18 Signature
      '',                             // 19 Schedule Conflict
      '',                             // 20 Supervisor Signature
      '',                             // 21 Send to Corp
      '',                             // 22 Denied2 Reason (W열)
      'Requested'                     // 23 Status (X열)
    ];
    Logger.log('rowData: ' + JSON.stringify(rowData));

    // 메인 시트 저장 (항상)
    mainSheet.appendRow(rowData);
    SpreadsheetApp.flush();
    Logger.log('Row appended to sheet');

    // === 헤더에서 컬럼 인덱스 동적 추출 ===
    const headersRow = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const headerIndex = {};
    headersRow.forEach((h, i) => headerIndex[h] = i);
    const lastRow = mainSheet.getLastRow();

    // Q열(Excuse Note): 무조건 텍스트만
    mainSheet.getRange(lastRow, headerIndex["Excuse Note"]+1).setValue(formData.excuseType);

    // R열(Excuse Note File): 하이퍼링크
    if (formData.excuseFileUrl && formData.excuseFileUrl.trim() !== '' && formData.excuseType === 'Submitted') {
      const excuseFileCell = mainSheet.getRange(lastRow, headerIndex["Excuse Note File"]+1);
      const fileName = formData.excuseFile || 'View File';
      const hyperlinkFormula = createHyperlink(formData.excuseFileUrl, fileName);
      if (hyperlinkFormula) {
        excuseFileCell.setFormula(hyperlinkFormula);
        Logger.log('Excuse note hyperlink set: ' + hyperlinkFormula);
      }
      // Q열을 다시 'Submitted'로 강제 (excuseType이 Submitted일 때만)
      mainSheet.getRange(lastRow, headerIndex["Excuse Note"]+1).setValue('Submitted');
    }

    // 디버깅 로그 추가
    Logger.log('Checking office/department save conditions:');
    Logger.log('sheetName: ' + sheetName);
    Logger.log('formData.office: ' + formData.office);
    Logger.log('formData.office !== "Taken": ' + (formData.office !== 'Taken'));
    Logger.log('Condition result: ' + ((sheetName === 'Incident Notice' || sheetName === 'Time-Off Request') && formData.office && formData.office !== 'Taken'));

    // Incident Notice, Time-Off Request만 부서/오피스별 시트에도 저장
    if ((sheetName === 'Incident Notice' || sheetName === 'Time-Off Request') && formData.office && formData.office !== 'Taken') {
      Logger.log('Trying to save to office/department spreadsheet: ' + formData.office);
      // Corporate 오피스 + 부서가 있는 경우: 부서별 시트
      let targetSpreadsheetId = null;
      if (formData.office === 'Corporate' && formData.department) {
        targetSpreadsheetId = OFFICE_SPREADSHEET_IDS[formData.department];
      } else {
        // 그 외 오피스: 오피스별 시트
        targetSpreadsheetId = OFFICE_SPREADSHEET_IDS[formData.office];
      }
      if (targetSpreadsheetId) {
        try {
          const targetSs = SpreadsheetApp.openById(targetSpreadsheetId);
          let targetSheet = targetSs.getSheetByName(sheetName);
          if (!targetSheet) {
            targetSheet = targetSs.insertSheet(sheetName);
            const headers = [
              "Type of Request", "Name", "Today's Date", "Email", "Office", "Department", "Position",
              "Type of Incident", "Cancelling Incident", "ETA", "ETD", "Start Date", "End Date",
              "Reason for Incident", "Opening", "Closing", "Excuse Note", "Excuse Note File",
              "Signature", "Schedule Conflict", "Supervisor Signature", "Send to Corp", "Status"
            ];
            targetSheet.appendRow(headers);
            Logger.log('Target sheet created and headers appended');
          }
          targetSheet.appendRow(rowData);
          SpreadsheetApp.flush();
          Logger.log('Data saved to office/department spreadsheet successfully');
          
          // 헤더에서 컬럼 인덱스 동적 추출
          const headersRow = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
          const headerIndex = {};
          headersRow.forEach((h, i) => headerIndex[h] = i);
          const lastRow = targetSheet.getLastRow();
          
          if (formData.excuseFileUrl && formData.excuseFileUrl.trim() !== '' && formData.excuseType === 'Submitted') {
            const excuseFileCell = targetSheet.getRange(lastRow, headerIndex["Excuse Note File"]+1);
            const fileName = formData.excuseFile || 'View File';
            const hyperlinkFormula = createHyperlink(formData.excuseFileUrl, fileName);
            if (hyperlinkFormula) {
              excuseFileCell.setFormula(hyperlinkFormula);
              Logger.log('Office/Department sheet excuse note hyperlink set: ' + hyperlinkFormula);
            }
            // Q열을 다시 'Submitted'로 강제 (excuseType이 Submitted일 때만)
            targetSheet.getRange(lastRow, headerIndex["Excuse Note"]+1).setValue('Submitted');
          }
        } catch (error) {
          Logger.log('Error saving to office/department spreadsheet: ' + error);
          throw error;
        }
      }
    }

    // PDF 첨부 이메일 발송 (D컬럼: email)
    const email = formData.email;
    Logger.log('Before PDF/email block, email: ' + email);
    if (email && email.indexOf('@') > -1) {
      Logger.log('Entering PDF/email block');
      try {
        const lastRow = mainSheet.getLastRow();
        Logger.log('lastRow for PDF: ' + lastRow);
        const pdfResult = generatePDFfromSheet(sheetName, lastRow);
        Logger.log('generatePDFfromSheet returned');
        const pdfBlob = pdfResult.pdfBlob;
        const emailBody =
          "Thank you for your submission.\n\n" +
          "Please find the attached PDF for your records.\n\n" +
          "You will receive an email with the approval result within a few days.\n\n" +
          "Note: If you don't receive an email within a week, please check your spam/junk folder.\n\n" +
          "Sincerely,\nSmileland Dental HR";
        try {
          MailApp.sendEmail({
            to: email,
            cc: getEmailAddresses(formData.office, formData.department),
            subject: 'Your Absence Request/Report Submission',
            body: emailBody,
            attachments: [pdfBlob],
            name: "Smileland Dental HR"
          });
          Logger.log('이메일 발송 성공');
        } catch (e) {
          Logger.log('이메일 발송 실패: ' + e);
          throw e;
        }
        Logger.log('Email sent to: ' + email);
      } catch (pdfEmailError) {
        Logger.log('Error in PDF generation or email sending: ' + pdfEmailError);
        throw pdfEmailError;
      }
    } else {
      Logger.log('Email block skipped, email: ' + email);
    }

    // === 캘린더 등록 추가 ===
    if (formData.incidentType) { // Cancelling 포함 모든 incidentType
      try {
        Logger.log('=== 캘린더 등록 시작 ===');
        Logger.log('formData.incidentType: ' + formData.incidentType);
        Logger.log('formData.office: ' + formData.office);
        Logger.log('formData.department: ' + formData.department);
        Logger.log('formData.startDate: ' + formData.startDate);
        
        let calendarId;
        let officeInitial; // 미리 선언
        
        // Cancelling 타입도 처음에는 requested 캘린더에 추가 (승인 후 approved - cancellation으로 이동)
        const targetStatus = 'requested';
        Logger.log('targetStatus: ' + targetStatus);
        
        if (formData.office === 'Corporate') {
          Logger.log('submitRequest department: ' + formData.department);
          calendarId = getCorporateCalendarId(formData.department, targetStatus);
          Logger.log('submitRequest calendarId: ' + calendarId);
          officeInitial = getOfficeInitial(formData.office, formData.department);
        } else {
          const statusCalendarIds = {
            'Ming': MING_STATUS_CALENDAR_IDS,
            'Bernard': BERNARD_STATUS_CALENDAR_IDS,
            'Delano': DELANO_STATUS_CALENDAR_IDS,
            'Tulare': TULARE_STATUS_CALENDAR_IDS,
            'Visalia': VISALIA_STATUS_CALENDAR_IDS,
            'Fresno': FRESNO_STATUS_CALENDAR_IDS,
            'California': CALIFORNIA_STATUS_CALENDAR_IDS,
            'Ortho': ORTHO_STATUS_CALENDAR_IDS,
            'Office Billing': OFFICE_BILLING_STATUS_CALENDAR_IDS
          };
          calendarId = statusCalendarIds[formData.office]?.[targetStatus];
          officeInitial = getOfficeInitial(formData.office, formData.department); // 추가
          if (formData.office === 'Visalia') {
            Logger.log('=== VISALIA OFFICE DETECTED ===');
            Logger.log('Visalia calendar IDs: ' + JSON.stringify(VISALIA_STATUS_CALENDAR_IDS));
            Logger.log('Target status: ' + targetStatus);
            Logger.log('Calendar ID: ' + calendarId);
          }
        }

        Logger.log('calendarId: ' + calendarId);
        Logger.log('formData.startDate: ' + formData.startDate);
        
        if (calendarId && formData.startDate) {
          Logger.log('캘린더 ID와 시작 날짜가 모두 존재함');
          const calendar = CalendarApp.getCalendarById(calendarId);
          Logger.log('calendar 객체: ' + (calendar ? 'found' : 'not found'));
          if (calendar) {
            const eventTitle = `${officeInitial}${formData.name}(${formData.incidentType})`;
            const eventDate = parseDateWithTimezone(formData.startDate);
            eventDate.setHours(0, 0, 0, 0);
            Logger.log('eventTitle: ' + eventTitle);
            Logger.log('eventDate: ' + eventDate);
            
            // Cancelling 타입인 경우 description 추가
            let eventOptions = {};
            if (formData.incidentType === 'Cancelling' && formData.cancelTargetType) {
              eventOptions.description = `Cancelling Incident: ${formData.cancelTargetType}`;
              Logger.log('eventOptions.description: ' + eventOptions.description);
            }
            
            // start date와 end date가 다른 경우 기간 이벤트로 생성
            const endDate = parseDateWithTimezone(formData.endDate);
            endDate.setHours(0, 0, 0, 0);
            
            if (eventDate.getTime() === endDate.getTime()) {
              // 같은 날짜인 경우 하루 이벤트
              calendar.createAllDayEvent(eventTitle, eventDate, eventOptions);
            } else {
              // 다른 날짜인 경우 기간 이벤트 (end date는 포함하지 않음)
              const endDateExclusive = new Date(endDate);
              endDateExclusive.setDate(endDateExclusive.getDate() + 1);
              calendar.createAllDayEvent(eventTitle, eventDate, endDateExclusive, eventOptions);
            }
            Logger.log('New event created in calendar: ' + calendarId + ' with status: ' + targetStatus);
          } else {
            Logger.log('캘린더를 찾을 수 없음: ' + calendarId);
          }
        } else {
          Logger.log('캘린더 ID 또는 시작 날짜가 없음');
          Logger.log('calendarId: ' + calendarId);
          Logger.log('formData.startDate: ' + formData.startDate);
        }
      } catch (err) {
        Logger.log('Calendar create error: ' + err);
        throw err;
      }
    }

    Logger.log('submitRequest finished successfully');
    return { success: true };
  } catch (error) {
    Logger.log('Error in submitRequest: ' + error);
    throw error;
  }
}

function generatePDFfromSheet(sheetName, rowNumber) {
  Logger.log('generatePDFfromSheet called, sheetName: ' + sheetName + ', rowNumber: ' + rowNumber);
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }
  Logger.log('Sheet: ' + (sheet ? 'found' : 'not found'));
  const data = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Data for PDF: ' + JSON.stringify(data));

  // 컬럼 라벨 (Signature 포함)
  const headers = [
    "Type of Request", "Name", "Today's Date", "Email", "Office", "Department", "Position",
    "Type of Incident", "Cancelling Incident", "ETA", "ETD", "Start Date", "End Date",
    "Reason for Incident", "Opening", "Closing", "Excuse Note", "Excuse Note File",
    "Signature", "Schedule Conflict", "Supervisor Signature", "Send to Corp", "Status"
  ];

  const doc = DocumentApp.create(sheetName === 'Incident Notice' ? 'Incident Notice PDF' : 
                                sheetName === 'Time-Off Request' ? 'Time-Off PDF' : 
                                sheetName === 'Taken' ? 'Taken PDF' : 'Document PDF');
  const body = doc.getBody();

  // === 제목/서브타이틀 동적 처리 ===
  if (sheetName === 'Incident Notice') {
    body.appendParagraph('Incident Notice (Requested in Advance)')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingAfter(6); // 간격 줄임
    body.appendParagraph('(Less than 30 days notice)')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingAfter(8); // 간격 줄임
  } else if (sheetName === 'Time-Off Request') {
    body.appendParagraph('Time-Off Request')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingAfter(6); // 간격 줄임
    body.appendParagraph('(30 days or more notice)')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingAfter(8); // 간격 줄임
  } else if (sheetName === 'Taken') {
    body.appendParagraph('Taken')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingAfter(8); // 간격 줄임
  } else {
    body.appendParagraph(sheetName)
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingAfter(8); // 간격 줄임
  }

  body.appendHorizontalRule();

  // === Type of Request(첫번째 컬럼) 제외하고 항목 출력 ===
  for (let i = 1; i < headers.length; i++) { // i=1부터 시작 (Type of Request 제외)
    let value = data[i];
    let safeValue = (value == null ? '' : String(value));
    // ETA, ETD는 'Tue Jun 03 2025 12:53:00' 형식으로 변환
    if (headers[i] === "ETA") {
      const etaValue = data[9];
      if (etaValue && /^\d{1,2}:\d{2}/.test(etaValue)) {
        safeValue = etaValue;
      } else if (etaValue && !isNaN(Date.parse(etaValue))) {
        const d = new Date(etaValue);
        const timeStr = d.toTimeString().slice(0, 5);
        safeValue = timeStr;
      } else {
        safeValue = 'N/A';
      }
    }
    if (headers[i] === "ETD") {
      const etdValue = data[10];
      if (etdValue && /^\d{1,2}:\d{2}/.test(etdValue)) {
        safeValue = etdValue;
      } else if (etdValue && !isNaN(Date.parse(etdValue))) {
        const d = new Date(etdValue);
        const timeStr = d.toTimeString().slice(0, 5);
        safeValue = timeStr;
      } else {
        safeValue = 'N/A';
      }
    }
    if (
      headers[i] === "Today's Date" ||
      headers[i] === "Start Date" ||
      headers[i] === "End Date"
    ) {
      if (safeValue && !isNaN(Date.parse(safeValue))) {
        safeValue = new Date(safeValue).toDateString();
      }
    }
    if (safeValue.trim() === '') safeValue = '';

    if (headers[i] === "Status" || 
        headers[i] === "Schedule Conflict" || 
        headers[i] === "Supervisor Signature" || 
        headers[i] === "Send to Corp") {
      continue;
    }
    if (headers[i] === "Excuse Note" && !safeValue) {
      continue;
    }
    if (headers[i] === "Excuse Note File" && (!safeValue || data[16] !== 'Submitted' || !safeValue.trim())) {
      continue;
    }

    if (headers[i] === "Signature") {
      if (safeValue.startsWith('data:image')) {
        body.appendParagraph('Signature:').editAsText().setBold(true);
        const imageBlob = Utilities.newBlob(Utilities.base64Decode(safeValue.split(',')[1]), 'image/png', 'signature.png');
        body.appendImage(imageBlob).setWidth(200); // 서명 이미지 크기 줄임
        body.appendParagraph('').setSpacingAfter(6); // 간격 줄임
      } else {
        const p = body.appendParagraph('Signature: N/A');
        p.editAsText().setBold(0, 8, true);
        p.setSpacingAfter(6); // 간격 줄임
      }
    } else {
      const label = String(headers[i]) + ': ';
      let displayValue = safeValue || 'N/A';
      
      // Excuse Note File의 경우 URL을 클릭 가능한 링크로 표시
      if (headers[i] === "Excuse Note File" && safeValue && safeValue.startsWith('http')) {
        displayValue = `[View File](${safeValue})`;
      }
      
      const p = body.appendParagraph(label + displayValue);
      p.editAsText().setBold(0, label.length - 1, true);
      p.setSpacingAfter(4); // 간격 더 줄임
    }
  }

  body.appendHorizontalRule();

  body.appendParagraph('Smileland Dental')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setFontSize(10)
      .setForegroundColor('#888888')
      .setSpacingAfter(0); // 하단 간격 제거

  doc.saveAndClose();
  const docFile = DriveApp.getFileById(doc.getId());
  Logger.log('Document created, docId: ' + doc.getId());
  const pdfBlob = docFile.getAs('application/pdf');
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  Logger.log('PDF blob created, folder found: ' + folder.getName());
  const pdfFile = folder.createFile(pdfBlob);
  // 파일명: Type of Request_Office_Name_Type of Incident_Start Date.pdf
  let startDate = data[11];
  if (startDate && !/^\d{4}-\d{2}-\d{2}$/.test(startDate)) {
    const d = new Date(startDate);
    if (!isNaN(d.getTime())) {
      startDate = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }
  }
  const fileName = `${data[0]}_${data[4]}_${data[1]}_${data[7]}_${startDate}.pdf`;
  pdfFile.setName(fileName);
  Logger.log('PDF file saved: ' + pdfFile.getUrl());
  return { fileUrl: pdfFile.getUrl(), pdfBlob: pdfBlob };
}

function onEdit(e) {
  Logger.log('onEdit 진입');
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const STATUS_COL = 24; // Y열 (Status)
    Logger.log('range.getColumn()=' + range.getColumn() + ', range.getRow()=' + range.getRow());
    // === 기존 Status 변경 로직 ===
    if (range.getColumn() === STATUS_COL && range.getRow() > 1) {
      Logger.log('Status 컬럼 조건 통과');
      const status = String(range.getValue() || '').toLowerCase().trim();
      Logger.log('status=' + status);
      // === status가 비었으면 아무 동작도 하지 않고 종료 ===
      if (!status) {
        Logger.log('Status가 비어있으므로 캘린더/이메일 동작 없이 종료');
        return;
      }
      const name = sheet.getRange(range.getRow(), 2).getValue();   // B열: Name
      const email = sheet.getRange(range.getRow(), 4).getValue();  // D열: Email
      Logger.log('email=' + email);
      const office = sheet.getRange(range.getRow(), 5).getValue(); // F열: Office
      const dateOfIncident = sheet.getRange(range.getRow(), 12).getValue(); // L열: Start Date (13번째)
      const attendanceType = sheet.getRange(range.getRow(), 8).getValue(); // I열: Type of Incident (9번째)
      const reason = sheet.getRange(range.getRow(), 14).getValue(); // N열: Reason for Incident (15번째)
      const cancelTargetType = sheet.getRange(range.getRow(), 9).getValue(); // J열: Cancelling Incident (10번째)

      const department = sheet.getRange(range.getRow(), 6).getValue();
      Logger.log('onEdit department: ' + department);

      // === Corporate/Ming/Bernard/Delano 오피스 상태 변경 시 캘린더 이동 ===
      if ((office === 'Corporate' || office === 'Ming' || office === 'Bernard' || office === 'Delano' || office === 'Tulare' || office === 'Visalia' || office === 'Fresno' || office === 'California' || office === 'Ortho' || office === 'Office Billing')) {
        const oldStatus = e.oldValue ? e.oldValue.toLowerCase() : '';
        const newStatus = status;
        const eventDate = parseDateWithTimezone(dateOfIncident);
        eventDate.setHours(0, 0, 0, 0);
        const officeInitial = getOfficeInitial(office, department);
        const attendanceEventTitle = `${officeInitial}${name}(${attendanceType})`;
        
        // 1. 모든 status 캘린더에서 이벤트 삭제
        let STATUS_CALENDAR_IDS;
        if (office === 'Corporate') {
          // 부서별 + 전체 Corporate 상태별 캘린더를 모두 합침
          STATUS_CALENDAR_IDS = {
            ...CORPORATE_STATUS_CALENDAR_IDS,
            ...(CORPORATE_DEPARTMENT_STATUS_CALENDAR_IDS[department] || {})
          };
          Logger.log('onEdit STATUS_CALENDAR_IDS: ' + JSON.stringify(STATUS_CALENDAR_IDS));
        } else if (office === 'Ming') {
          STATUS_CALENDAR_IDS = MING_STATUS_CALENDAR_IDS;
        } else if (office === 'Bernard') {
          STATUS_CALENDAR_IDS = BERNARD_STATUS_CALENDAR_IDS;
        } else if (office === 'Delano') {
          STATUS_CALENDAR_IDS = DELANO_STATUS_CALENDAR_IDS;
        } else if (office === 'Tulare') {
          STATUS_CALENDAR_IDS = TULARE_STATUS_CALENDAR_IDS;
        } else if (office === 'Visalia') {
          STATUS_CALENDAR_IDS = VISALIA_STATUS_CALENDAR_IDS;
          Logger.log('Visalia office detected - using VISALIA_STATUS_CALENDAR_IDS');
          Logger.log('Visalia calendar IDs: ' + JSON.stringify(VISALIA_STATUS_CALENDAR_IDS));
        } else if (office === 'Fresno') {
          STATUS_CALENDAR_IDS = FRESNO_STATUS_CALENDAR_IDS;
        } else if (office === 'California') {
          STATUS_CALENDAR_IDS = CALIFORNIA_STATUS_CALENDAR_IDS;
        } else if (office === 'Ortho') {
          STATUS_CALENDAR_IDS = ORTHO_STATUS_CALENDAR_IDS;
        } else if (office === 'Office Billing') {
          STATUS_CALENDAR_IDS = OFFICE_BILLING_STATUS_CALENDAR_IDS;
        }

        // 모든 상태의 캘린더에서 이벤트 삭제
        for (let key in STATUS_CALENDAR_IDS) {
          const calendarId = STATUS_CALENDAR_IDS[key];
          if (calendarId) {
            const calendar = CalendarApp.getCalendarById(calendarId);
            if (calendar) {
              Logger.log(`Checking calendar: ${calendarId} for event: ${attendanceEventTitle}`);
              // 더 넓은 범위로 이벤트 검색 (전후 1일 포함)
              const searchStartDate = new Date(eventDate);
              searchStartDate.setDate(searchStartDate.getDate() - 1);
              const searchEndDate = new Date(eventDate);
              searchEndDate.setDate(searchEndDate.getDate() + 1);
              const events = calendar.getEvents(searchStartDate, searchEndDate);
              Logger.log(`Found ${events.length} events for date: ${eventDate}`);
              for (let i = 0; i < events.length; i++) {
                Logger.log(`Event ${i}: ${events[i].getTitle()}`);
                // 제목이 정확히 일치하거나 부분적으로 일치하는 경우 삭제
                const eventTitle = events[i].getTitle();
                if (eventTitle === attendanceEventTitle || 
                    eventTitle.includes(name) && eventTitle.includes(attendanceType)) {
                  events[i].deleteEvent();
                  Logger.log('Deleted event from calendar: ' + calendarId);
                  Logger.log('Event title matched and deleted: ' + attendanceEventTitle);
                  Logger.log('Actually deleted event: ' + eventTitle);
                }
              }
            }
          }
        }

        // Visalia 특화 추가 안전장치
        if (office === 'Visalia') {
          Logger.log('=== VISALIA SPECIAL CLEANUP ===');
          for (let key in STATUS_CALENDAR_IDS) {
            const calendarId = STATUS_CALENDAR_IDS[key];
            if (calendarId) {
              const calendar = CalendarApp.getCalendarById(calendarId);
              if (calendar) {
                // 해당 날짜의 모든 이벤트를 다시 확인
                const events = calendar.getEventsForDay(eventDate);
                Logger.log(`Visalia cleanup - Calendar ${calendarId}: found ${events.length} events`);
                for (let i = events.length - 1; i >= 0; i--) { // 역순으로 삭제
                  const eventTitle = events[i].getTitle();
                  if (eventTitle.includes(name) || eventTitle.includes(attendanceType)) {
                    events[i].deleteEvent();
                    Logger.log('Visalia cleanup: Deleted event: ' + eventTitle);
                  }
                }
              }
            }
          }
        }

        // 2. 새 status가 있는 경우에만 새 캘린더에 이벤트 생성
        if (newStatus && ALLOWED_STATUSES.includes(newStatus)) {
          let targetCalId;
          
          // Cancelling 타입이고 approved - cancellation 상태인 경우 cancelled 캘린더에 추가
          if (attendanceType === 'Cancelling' && newStatus === 'approved - cancellation') {
            targetCalId = STATUS_CALENDAR_IDS['cancelled'];
            Logger.log(`Cancelling event - using cancelled calendar: ${targetCalId}`);
          } else {
            targetCalId = STATUS_CALENDAR_IDS[newStatus];
            Logger.log(`Regular event - using ${newStatus} calendar: ${targetCalId}`);
          }
          
          Logger.log(`Creating new event in calendar: ${targetCalId} for status: ${newStatus}`);
          const newCal = CalendarApp.getCalendarById(targetCalId);
          if (newCal) {
            // Cancelling 타입이면 description에 Cancelling Incident 값 추가
            let eventOptions = {};
            if (attendanceType === 'Cancelling' && cancelTargetType) {
              eventOptions.description = `Cancelling Incident: ${cancelTargetType}`;
            }
            // start date와 end date가 다른 경우 기간 이벤트로 생성
            const endDate = parseDateWithTimezone(sheet.getRange(range.getRow(), 13).getValue()); // N열: End Date
            endDate.setHours(0, 0, 0, 0);
            
            if (eventDate.getTime() === endDate.getTime()) {
              // 같은 날짜인 경우 하루 이벤트
              newCal.createAllDayEvent(attendanceEventTitle, eventDate, eventOptions);
            } else {
              // 다른 날짜인 경우 기간 이벤트 (end date는 포함하지 않음)
              const endDateExclusive = new Date(endDate);
              endDateExclusive.setDate(endDateExclusive.getDate() + 1);
              newCal.createAllDayEvent(attendanceEventTitle, eventDate, endDateExclusive, eventOptions);
            }
            Logger.log('Created event in new calendar: ' + attendanceEventTitle);
            Logger.log('New event created successfully in: ' + targetCalId);
          }
        }

        // === Cancelling 처리: 기존 이벤트 찾아서 삭제 ===
        if (attendanceType === 'Cancelling') {
          // Cancelling Incident로 지정된 기존 이벤트 삭제
          const mainData = sheet.getDataRange().getValues();
          let foundRow = null;
          let foundRowIdx = -1;
          
          // 취소하려는 기간
          const cancelStartDate = parseDateWithTimezone(dateOfIncident);
          const cancelEndDate = parseDateWithTimezone(sheet.getRange(range.getRow(), 13).getValue()); // N열: End Date
          cancelStartDate.setHours(0, 0, 0, 0);
          cancelEndDate.setHours(0, 0, 0, 0);
          
          for (let i = 1; i < mainData.length; i++) { // i=1부터: 헤더 제외
            const rowName = mainData[i][1];
            const rowStartDate = mainData[i][11];
            const rowEndDate = mainData[i][12];
            const rowType = mainData[i][7];
            
            if (
              rowName === name &&
              rowType === cancelTargetType &&
              rowType !== 'Cancelling'
            ) {
              const originalStartDate = parseDateWithTimezone(rowStartDate);
              const originalEndDate = parseDateWithTimezone(rowEndDate);
              originalStartDate.setHours(0, 0, 0, 0);
              originalEndDate.setHours(0, 0, 0, 0);
              
              // 기간이 겹치는지 확인
              if (cancelStartDate <= originalEndDate && cancelEndDate >= originalStartDate) {
                foundRow = mainData[i];
                foundRowIdx = i;
                break;
              }
            }
          }
          
          if (foundRow) {
            // 모든 상태의 캘린더에서 해당 이벤트 처리
            for (let key in STATUS_CALENDAR_IDS) {
              const calendarId = STATUS_CALENDAR_IDS[key];
              if (calendarId) {
                const calendar = CalendarApp.getCalendarById(calendarId);
                if (calendar) {
                  const officeInitialToDelete = getOfficeInitial(office, department);
                  const eventTitleToDelete = `${officeInitialToDelete}${name}(${cancelTargetType})`;
                  
                  // 원본 이벤트의 기간
                  const originalStartDate = parseDateWithTimezone(foundRow[11]);
                  const originalEndDate = parseDateWithTimezone(foundRow[12]);
                  originalStartDate.setHours(0, 0, 0, 0);
                  originalEndDate.setHours(0, 0, 0, 0);
                  
                  // 더 넓은 범위로 이벤트 검색
                  const searchStartDate = new Date(originalStartDate);
                  searchStartDate.setDate(searchStartDate.getDate() - 1);
                  const searchEndDate = new Date(originalEndDate);
                  searchEndDate.setDate(searchEndDate.getDate() + 1);
                  
                  const events = calendar.getEvents(searchStartDate, searchEndDate);
                  for (let i = 0; i < events.length; i++) {
                    if (events[i].getTitle() === eventTitleToDelete) {
                      const event = events[i];
                      const eventStart = event.getStartTime();
                      const eventEnd = event.getEndTime();
                      
                      // 이벤트 삭제
                      event.deleteEvent();
                      Logger.log('Deleted original event: ' + eventTitleToDelete + ' from calendar: ' + calendarId);
                      
                      // 부분 삭제 후 남은 기간이 있는지 확인하고 새 이벤트 생성
                      let remainingStartDate = null;
                      let remainingEndDate = null;
                      
                      // 취소 시작일 이전에 남은 기간이 있는지
                      if (originalStartDate < cancelStartDate) {
                        remainingStartDate = originalStartDate;
                        remainingEndDate = new Date(cancelStartDate);
                        remainingEndDate.setDate(remainingEndDate.getDate() - 1);
                      }
                      
                      // 취소 종료일 이후에 남은 기간이 있는지
                      if (originalEndDate > cancelEndDate) {
                        if (remainingStartDate === null) {
                          remainingStartDate = new Date(cancelEndDate);
                          remainingStartDate.setDate(remainingStartDate.getDate() + 1);
                        }
                        remainingEndDate = originalEndDate;
                      }
                      
                      // 남은 기간이 있으면 새 이벤트 생성
                      if (remainingStartDate && remainingEndDate && remainingStartDate <= remainingEndDate) {
                        const remainingEventTitle = `${officeInitialToDelete}${name}(${cancelTargetType})`;
                        const remainingEndDateExclusive = new Date(remainingEndDate);
                        remainingEndDateExclusive.setDate(remainingEndDateExclusive.getDate() + 1);
                        
                        calendar.createAllDayEvent(remainingEventTitle, remainingStartDate, remainingEndDateExclusive);
                        Logger.log('Created remaining event: ' + remainingEventTitle + ' from ' + remainingStartDate.toDateString() + ' to ' + remainingEndDate.toDateString());
                      }
                      
                      break; // 해당 이벤트를 찾았으므로 루프 종료
                    }
                  }
                }
              }
            }
          }
        }
        
        // Y열 변경 시에는 캘린더 처리만 완료하고 종료 (이메일 발송 없음)
        Logger.log('Y열 Status 변경 - 캘린더 처리 완료, 이메일 발송 없음');
      }
    }
    
    // === Email to Employee 컬럼(Z열) 체크 시 이메일 발송 ===
    const EMAIL_TO_EMPLOYEE_COL = 25; // Z열 (Email to Employee)
    if (range.getColumn() === EMAIL_TO_EMPLOYEE_COL && range.getRow() > 1) {
      Logger.log('Email to Employee 컬럼 조건 통과');
      const emailToEmployee = range.getValue();
      Logger.log('emailToEmployee=' + emailToEmployee);
      
      // 체크가 들어온 경우에만 이메일 발송
      if (emailToEmployee === true || emailToEmployee === 'TRUE' || emailToEmployee === 'true') {
        Logger.log('Email to Employee 체크됨 - 이메일 발송 시작');
        
        const name = sheet.getRange(range.getRow(), 2).getValue();   // B열: Name
        const email = sheet.getRange(range.getRow(), 4).getValue();  // D열: Email
        const status = String(sheet.getRange(range.getRow(), STATUS_COL).getValue() || '').toLowerCase().trim();
        const cancelTargetType = sheet.getRange(range.getRow(), 9).getValue(); // J열: Cancelling Incident (10번째)
        
        Logger.log('email=' + email + ', status=' + status);
        
        if (email && status && ALLOWED_STATUSES.includes(status)) {
          let subject, bodyText;
          let shouldSendEmail = true;
          
          switch (status) {
            case 'denied':
              subject = `Absence Request/Report Denied`;
              const deniedReason = sheet.getRange(range.getRow(), 23).getValue(); // X열
              bodyText = `${name},\n\nWe regret to inform you that your request has been denied.`;
              if (deniedReason && String(deniedReason).trim() !== '') {
                bodyText += `\n${deniedReason}`;
              }
              // 제출 정보 추가
              const submittedDate = sheet.getRange(range.getRow(), 3).getValue(); // Today's Date (C열)
              const typeOfRequest = sheet.getRange(range.getRow(), 1).getValue(); // Type of Request (A열)
              const typeOfIncident = sheet.getRange(range.getRow(), 8).getValue(); // Type of Incident (I열)
              const startDate = sheet.getRange(range.getRow(), 12).getValue(); // Start Date (M열)
              const endDate = sheet.getRange(range.getRow(), 13).getValue(); // End Date (N열)
              bodyText += `\n\nSubmitted Date: ${formatDateForEmail(submittedDate)}`;
              bodyText += `\nType of Request: ${typeOfRequest}`;
              bodyText += `\nType of Incident: ${typeOfIncident}`;
              if (typeOfIncident === 'Cancelling' && cancelTargetType) {
                bodyText += `\nCancelling Incident: ${cancelTargetType}`;
              }
              bodyText += `\nRequested Date: ${formatDateForEmail(startDate)} ~ ${formatDateForEmail(endDate)}`;
              bodyText += `\n\nThank you so much.\nSincerely,\nSmileland Dental HR`;
              break;
            case 'approved - cancellation':
              subject = `Absence Request/Report Cancelled`;
              bodyText = `${name},\n\nYour incident notice request has been cancelled.`;
              // 제출 정보 추가
              const submittedDate2 = sheet.getRange(range.getRow(), 3).getValue();
              const typeOfRequest2 = sheet.getRange(range.getRow(), 1).getValue();
              const typeOfIncident2 = sheet.getRange(range.getRow(), 8).getValue();
              const startDate2 = sheet.getRange(range.getRow(), 12).getValue();
              const endDate2 = sheet.getRange(range.getRow(), 13).getValue();
              bodyText += `\n\nSubmitted Date: ${formatDateForEmail(submittedDate2)}`;
              bodyText += `\nType of Request: ${typeOfRequest2}`;
              bodyText += `\nType of Incident: ${typeOfIncident2}`;
              if (typeOfIncident2 === 'Cancelling' && cancelTargetType) {
                bodyText += `\nCancelling Incident: ${cancelTargetType}`;
              }
              bodyText += `\nRequested Date: ${formatDateForEmail(startDate2)} ~ ${formatDateForEmail(endDate2)}`;
              bodyText += `\n\nThank you so much.\nSincerely,\nSmileland Dental HR`;
              break;
            case 'pending':
              subject = `Absence Request/Report Pending`;
              bodyText = `${name},\n\nYour incident notice request is pending review.`;
              // 제출 정보 추가
              const submittedDate3 = sheet.getRange(range.getRow(), 3).getValue();
              const typeOfRequest3 = sheet.getRange(range.getRow(), 1).getValue();
              const typeOfIncident3 = sheet.getRange(range.getRow(), 8).getValue();
              const startDate3 = sheet.getRange(range.getRow(), 12).getValue();
              const endDate3 = sheet.getRange(range.getRow(), 13).getValue();
              bodyText += `\n\nSubmitted Date: ${formatDateForEmail(submittedDate3)}`;
              bodyText += `\nType of Request: ${typeOfRequest3}`;
              bodyText += `\nType of Incident: ${typeOfIncident3}`;
              if (typeOfIncident3 === 'Cancelling' && cancelTargetType) {
                bodyText += `\nCancelling Incident: ${cancelTargetType}`;
              }
              bodyText += `\nRequested Date: ${formatDateForEmail(startDate3)} ~ ${formatDateForEmail(endDate3)}`;
              bodyText += `\n\nThank you so much.\nSincerely,\nSmileland Dental HR`;
              break;
            case 'approved':
              subject = `Absence Request/Report Approved`;
              bodyText = `${name},\n\nYour absence request/report has been approved.`;
              // 제출 정보 추가
              const submittedDate4 = sheet.getRange(range.getRow(), 3).getValue();
              const typeOfRequest4 = sheet.getRange(range.getRow(), 1).getValue();
              const typeOfIncident4 = sheet.getRange(range.getRow(), 8).getValue();
              const startDate4 = sheet.getRange(range.getRow(), 12).getValue();
              const endDate4 = sheet.getRange(range.getRow(), 13).getValue();
              bodyText += `\n\nSubmitted Date: ${formatDateForEmail(submittedDate4)}`;
              bodyText += `\nType of Request: ${typeOfRequest4}`;
              bodyText += `\nType of Incident: ${typeOfIncident4}`;
              if (typeOfIncident4 === 'Cancelling' && cancelTargetType) {
                bodyText += `\nCancelling Incident: ${cancelTargetType}`;
              }
              bodyText += `\nRequested Date: ${formatDateForEmail(startDate4)} ~ ${formatDateForEmail(endDate4)}`;
              bodyText += `\n\nThank you so much.\nSincerely,\nSmileland Dental HR`;
              break;
            case 'approved – excuse note required':
              subject = `Absence Request/Report Approved`;
              bodyText = `${name},\n\nYour absence request/report has been approved.\n\nPlease submit excuse note upon arrival.`;
              // 제출 정보 추가
              const submittedDate5 = sheet.getRange(range.getRow(), 3).getValue();
              const typeOfRequest5 = sheet.getRange(range.getRow(), 1).getValue();
              const typeOfIncident5 = sheet.getRange(range.getRow(), 8).getValue();
              const startDate5 = sheet.getRange(range.getRow(), 12).getValue();
              const endDate5 = sheet.getRange(range.getRow(), 13).getValue();
              bodyText += `\n\nSubmitted Date: ${formatDateForEmail(submittedDate5)}`;
              bodyText += `\nType of Request: ${typeOfRequest5}`;
              bodyText += `\nType of Incident: ${typeOfIncident5}`;
              if (typeOfIncident5 === 'Cancelling' && cancelTargetType) {
                bodyText += `\nCancelling Incident: ${cancelTargetType}`;
              }
              bodyText += `\nRequested Date: ${formatDateForEmail(startDate5)} ~ ${formatDateForEmail(endDate5)}`;
              bodyText += `\n\nThank you so much.\nSincerely,\nSmileland Dental HR`;
              break;
            default:
              shouldSendEmail = false;
          }
          
          Logger.log('shouldSendEmail=' + shouldSendEmail);
          try {
            if (shouldSendEmail && email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
              Logger.log('이메일 발송 시도');
              MailApp.sendEmail({
                to: email,
                cc: getEmailAddressesForStatusChange(sheet.getRange(range.getRow(), 5).getValue(), sheet.getRange(range.getRow(), 6).getValue()), // Office, Department (attendance@smlnd.com 제외)
                subject: subject,
                body: bodyText,
                name: "Smileland Dental HR"
              });
              Logger.log('이메일 발송 성공: ' + email);
              
              // 이메일 발송 성공 시 AA컬럼(26번째 컬럼)에 "Yes" 입력
              const emailSentCell = sheet.getRange(range.getRow(), 26); // AA열
              emailSentCell.setValue('Yes');
              Logger.log('AA컬럼에 "Yes" 입력 완료');
            } else {
              Logger.log('이메일 발송 조건 불충족');
            }
          } catch (mailError) {
            Logger.log('MailApp.sendEmail error: ' + mailError);
          }
        } else {
          Logger.log('이메일 발송 조건 불충족: email=' + email + ', status=' + status);
        }
      }
    }
  } catch (error) {
    Logger.log('Error in onEdit: ' + error);
  }
}

// 오피스 이니셜 생성 함수 추가
function getOfficeInitial(office, department) {
  if (office === 'California') return '(CA)';
  if (office === 'Corporate') {
    if (department === 'AR') return '(AR)';
    if (department === 'Call Center') return '(CC)';
    if (department === 'IT') return '(IT)';
    if (department === 'Marcom') return '(MC)';
    if (department === 'EXEC') return '(EXEC)';
    return '(C)';
  }
  if (office === 'Ortho') return '(O)';
  if (office === 'Office Billing') return '(OB)';
  return `(${office[0]})`;
}

// === Web App 파일 업로드 핸들러 ===
function doPost(e) {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    if (e && e.postData && e.postData.type.indexOf('multipart') === 0) {
      var parts = Utilities.parseMultipart(e.postData, e.postData.type);
      Logger.log('parts keys: ' + Object.keys(parts)); // 진단용 로그 추가
      var fileBlob = parts.file;
      if (!fileBlob) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'No file part found.' })).setMimeType(ContentService.MimeType.JSON);
      }
      var file = folder.createFile(fileBlob);
      return ContentService.createTextOutput(JSON.stringify({ success: true, fileUrl: file.getUrl() })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'No file uploaded.' })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// === Base64 첨부파일 업로드 핸들러 ===
function uploadBase64File(base64, fileName) {
  try {
    var content = base64.split(',')[1];
    var mimeType = base64.match(/^data:(.*);base64,/)[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(content), mimeType, fileName);
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var file = folder.createFile(blob);
    return { success: true, fileUrl: file.getUrl() };
  } catch (e) {
    return { success: false, message: '파일 업로드 실패: ' + e.message };
  }
}

// 날짜를 YYYY-MM-DD 형식으로 변환하는 함수
function formatDateForEmail(dateValue) {
  if (!dateValue) return '';
  if (typeof dateValue === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateValue)) {
    return dateValue;
  }
  const date = new Date(dateValue);
  if (isNaN(date.getTime())) return '';
  return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
}

// 하이퍼링크 생성 함수
function createHyperlink(url, text) {
  if (!url || url.trim() === '') return '';
  
  // URL이 이미 http로 시작하는지 확인
  if (url.startsWith('http')) {
    return `=HYPERLINK("${url}","${text || 'View File'}")`;
  }
  
  return url; // URL이 아니면 그대로 반환
}