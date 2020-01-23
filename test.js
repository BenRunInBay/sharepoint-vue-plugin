// Import the mount() method from the test utils
// and the component you want to test
import { mount } from "@vue/test-utils";
import * as SharePointConnector from "./project-template/src/lib/SharePointConnector.VuePlugin";

describe("SPC tests", () => {
  it("Can be installed", () => {
    expect(typeof SharePointConnector.default.install).toBe("function");
  });

  it("The Class instantiates and accepts a base URL", () => {
    let spc = new SharePointConnector.SharePointConnector("/base");
    expect(spc.baseUrl).toBe("/base");
  });

  it("Is in dev mode", () => {
    let spc = new SharePointConnector.SharePointConnector();
    expect(spc.inProduction).toBe(false);
  });

  it("Request digest from server", () => {
    let spc = new SharePointConnector.SharePointConnector(
      "https://sites.ey.com/sites/gmsa/testing/"
    );
    spc.getRequestDigest(() => {
      expect(spc.digestValue).toBeTruthy();
    });
  });
  it("Can read from GMSA Testing IssueTracking list", () => {
    let spc = new SharePointConnector.SharePointConnector(
      "https://sites.ey.com/sites/gmsa/testing/"
    );
    spc.getList({
      listName: "Issue Tracking",
      select: "ID,Title",
      top: 1,
      success(data) {
        expect(data.length).toBeGreaterThan(0);
      },
      failure(error) {
        expect(error).toBeDefined();
      },
      devStaticDataUrl: "TestData/issue-tracking.json"
    });
  });

  it("CAML Query test", () => {
    let spc = new SharePointConnector.SharePointConnector(
      "https://sites.ey.com/sites/gmsa/testing/"
    );
    spc.getList({
      listName: "Issue Tracking",
      queryXml: "",
      success(data) {
        expect(data.length).toBeGreaterThan(0);
      },
      failure(error) {
        expect(error).toBeDefined();
      },
      devStaticDataUrl: "TestData/issue-tracking.json"
    });
  });

  it("Add item to GMSA Testing IssueTracking list", () => {
    let spc = new SharePointConnector.SharePointConnector(
      "https://sites.ey.com/sites/gmsa/testing/"
    );
    spc.getRequestDigest(() => {
      spc.addListItem({
        listName: "Issue Tracking",
        itemData: {
          Title: "Test from SPC unit tester"
        },
        success(data, itemUrl) {
          expect(itemUrl.length).toBeTruthy();
        },
        failure(error) {
          expect(error).toBeDefined();
        }
      });
    });
  });

  it("Update item in GMSA Testing IssueTracking list", () => {
    let spc = new SharePointConnector.SharePointConnector(
      "https://sites.ey.com/sites/gmsa/testing/"
    );
    spc.getRequestDigest(() => {
      spc.updateListItem({
        listName: "Issue Tracking",
        itemUrl:
          "https://sites.ey.com/sites/gmsa/testing/_api/Web/Lists(guid'9cf26cf9-0e7f-4dc2-893f-b2e3779fe475')/Items(1)",
        itemData: {
          Title: "UPDATE Test from SPC unit tester"
        },
        success(data) {
          expect(data).toBeTruthy();
        },
        failure(error) {
          expect(error).toBeDefined();
        }
      });
    });
  });

  it("retrievePeopleProfile", () => {
    let spc = new SharePointConnector.SharePointConnector(
      "https://sites.ey.com/sites/gmsa/testing/"
    );
    spc.getRequestDigest(() => {
      spc.retrievePeopleProfile(
        "i:0#.f|membership|name@domain",
        null,
        data => {
          expect(data).toBeTruthy();
        },
        error => {
          expect(error).toBeDefined();
        }
      );
    });
  });

  it("retrieveCurrentUserProfile", () => {
    let spc = new SharePointConnector.SharePointConnector(
      "https://sites.ey.com/sites/gmsa/testing/"
    );
    spc.getRequestDigest(() => {
      spc.retrieveCurrentUserProfile(
        data => {
          expect(data).toBeTruthy();
        },
        error => {
          expect(error).toBeDefined();
        }
      );
    });
  });

  it("ensureSiteUserId", () => {
    let spc = new SharePointConnector.SharePointConnector(
      "https://sites.ey.com/sites/gmsa/testing/"
    );
    spc.getRequestDigest(() => {
      spc.ensureSiteUserId({
        accountName: "i:0#.f|membership|name@domain",
        success: function(data) {
          expect(data).toBeTruthy();
        },
        failure: function(error) {
          expect(error).toBeDefined();
        }
      });
    });
  });

  it("sendEmail", () => {
    let spc = new SharePointConnector.SharePointConnector(
      "https://sites.ey.com/sites/gmsa/testing/"
    );
    spc.getRequestDigest(() => {
      spc.sendEmail({
        from: "benjamin.hoffmann@ey.com",
        to: ["hoffmbe@gmail.com"],
        subject: "test",
        bodyHtml: "test body",
        success() {
          expect(true).toBe(true);
        },
        failure(error) {
          expect(error).toBeDefined();
        }
      });
    });
  });
});
