package com.heroku.java;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.annotation.JsonIgnore;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.sql.Timestamp;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;


public class Candidate {
    //    @Id
//    @GeneratedValue(strategy = GenerationType.IDENTITY)
//    @Column
    @Getter
    @Setter
    private long id;
    @Getter
    @Setter
    private long createId;
    @Setter
    @Getter
    private String name;
    @Setter
    @Getter
    private String nameKana;
    @Setter
    @Getter
    private String address;
    @Setter
    @Getter
    private String graduateSchool;
    @Setter
    @Getter
    private String note;
    @Setter
    @Getter
    private String compatibleField;
    @Setter
    @Getter
//    @Column(length = 5000)
    private String selfPR;
    @Setter
    @Getter
//    @Column(length = 1000)
    private String reference;
    @Setter
    @Getter
    @JsonFormat(pattern = "yyyy/MM/dd")
    private Timestamp birthDay;
    @Setter
    @Getter
    private String birthPlace;
    @Setter
    @Getter
    private boolean gender;
    @Setter
    @Getter
    // statusからworkStatusに変更する
    private String workStatus;
    @Setter
    @Getter
    //　候補者のスキルシートを表示・非表示する
    private int candidateStatus;
    @Setter
    @Getter
    private int experienceYears;
    @Setter
    @Getter
    private long unitPrice = 450000;
    @Setter
    @Getter
    private String experienceYearsInJapan;

//    @Setter
//    @Getter
//    @Column(columnDefinition="TEXT")
//    private String image;
//    @Setter
//    @Getter
//    @ManyToOne(optional = true, fetch = FetchType.EAGER)
//    private EnglishLevel englishLevel;

//    @Setter
//    @Getter
//    @ManyToOne(optional = true, fetch = FetchType.EAGER)
//    private JapaneseLevel japaneseLevel;
//    @Setter
//    @Getter
//    @ManyToOne(optional = true, fetch = FetchType.EAGER)
//    private WorkRole workRole;

//    @Setter
//    @Getter
//    @OneToOne(fetch = FetchType.EAGER)
//    private User user;

//    @Setter
//    @Getter
//    @OneToMany(mappedBy="candidate")
//    private Set<Project> projects;
//
//    @Setter
//    @Getter
//    @ManyToOne(optional = true, fetch = FetchType.EAGER)
//    private Customer customer;

//    @Setter
//    @Getter
//    @JsonIgnore
//    @OneToMany(mappedBy = "candidate")
//    private Set<CandidateSkill> skills = new HashSet<>();

    @JsonIgnore
    @Setter
    @Getter
    private Timestamp createAt;

    @JsonIgnore
    @Setter
    @Getter
    private Timestamp updatedAt;

    @JsonIgnore
    @Setter
    @Getter
    private boolean delete;

    //　英語と日本語を区別するため
    @Getter
    @Setter
    private Boolean sheetLanguage;
    //候補者の組織名
    @Getter
    @Setter
    private String organization;


}
